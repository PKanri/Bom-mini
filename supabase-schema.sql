-- Authenticated mode. Safe to rerun: existing business data is preserved.
begin;
create table if not exists public.bom_members (
  user_id uuid primary key references auth.users(id) on delete cascade,
  role text not null check (role in ('admin','branch')),
  branch text,
  check (role = 'admin' or (branch is not null and branch <> ''))
);
alter table public.bom_members enable row level security;
revoke all on public.bom_members from anon, authenticated;
grant select on public.bom_members to authenticated;
drop policy if exists own_membership on public.bom_members;
create policy own_membership on public.bom_members for select to authenticated
using (user_id = (select auth.uid()));

-- A separate row per monthly source/branch and per count date/branch.
-- JSON preserves the existing recipe and import formats without losing fields.
create table if not exists public.bom_documents (
  id text primary key,
  kind text not null check (kind in ('master','SALES','STOCK_IN','STOCK_OUT','EXTRA_IN','count')),
  branch text not null default '',
  period text not null default '',
  payload jsonb not null,
  revision bigint not null default 1 check (revision > 0),
  updated_at timestamptz not null default now(),
  check (jsonb_typeof(payload) = 'object'),
  check ((kind='master' and id='master' and branch='' and period='') or
    (kind<>'master' and branch<>'' and id=kind||':'||branch||':'||period)),
  check (kind='master' or (kind='count' and period ~ '^\d{4}-\d{2}-\d{2}$') or
    (kind in ('SALES','STOCK_IN','STOCK_OUT','EXTRA_IN') and period ~ '^\d{4}-\d{2}$'))
);
alter table public.bom_documents enable row level security;
revoke all on public.bom_documents from anon, authenticated;
grant select on public.bom_documents to authenticated;
drop policy if exists shared_read on public.bom_documents;
drop policy if exists member_read on public.bom_documents;
create policy member_read on public.bom_documents for select to authenticated using (
  exists (select 1 from public.bom_members m where m.user_id=(select auth.uid())
    and (m.role='admin' or bom_documents.kind='master' or bom_documents.branch=m.branch))
);

-- Only this function can write. It checks membership, scope and revisions.
-- The transaction lock prevents two clients from overwriting the same revision.
create or replace function public.bom_save(changes jsonb)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare
  member public.bom_members%rowtype;
  item jsonb;
  old public.bom_documents%rowtype;
  result jsonb := '[]'::jsonb;
  next_revision bigint;
begin
  select * into member from public.bom_members where user_id=auth.uid();
  if not found then raise exception 'No BOM membership' using errcode='42501'; end if;
  if jsonb_typeof(changes) <> 'array' or jsonb_array_length(changes)>10000 then
    raise exception 'Invalid changes';
  end if;
  perform pg_advisory_xact_lock(817241591);
  for item in select value from jsonb_array_elements(changes) loop
    if member.role <> 'admin' and not (
      (coalesce(item->>'kind','')='count' and coalesce(item->>'branch','')=member.branch)
      or (coalesce(item->>'kind','')='master' and coalesce(item->>'id','')='master'
          and coalesce(item->>'branch','')='' and coalesce(item->>'period','')=''
          and not coalesce((item->>'remove')::boolean,false))
    ) then
      raise exception 'This account can only edit shared BOM or its own branch counts' using errcode='42501';
    end if;
    select * into old from public.bom_documents where id=item->>'id';
    if found then
      if old.kind <> item->>'kind' or old.branch <> item->>'branch' or old.period <> item->>'period' then
        raise exception 'Document identity cannot change';
      end if;
      if (item->>'expected')::bigint is distinct from old.revision then
        raise exception 'CONFLICT: reload before saving' using errcode='40001';
      end if;
      next_revision := old.revision+1;
    else
      if item->>'expected' is not null then raise exception 'CONFLICT: document removed' using errcode='40001'; end if;
      next_revision := 1;
    end if;
    if coalesce((item->>'remove')::boolean,false) then
      delete from public.bom_documents where id=item->>'id';
    else
      insert into public.bom_documents(id,kind,branch,period,payload,revision)
      values(item->>'id',item->>'kind',item->>'branch',item->>'period',item->'payload',next_revision)
      on conflict(id) do update set payload=excluded.payload,revision=excluded.revision,updated_at=now();
    end if;
    result := result || jsonb_build_array(jsonb_build_object('id',item->>'id','revision',next_revision));
  end loop;
  return result;
end $$;
revoke all on function public.bom_save(jsonb) from public, anon;
grant execute on function public.bom_save(jsonb) to authenticated;
commit;

-- Allow members to check imported movement-file date coverage only.
-- No transaction rows, quantities or other branch identifiers are returned.
begin;
create or replace function public.bom_import_coverage()
returns table(kind text, period text, start_date text, end_date text, quantity_field text)
language sql stable security definer set search_path = '' as $$
  select distinct d.kind, d.period, d.payload->>'start', d.payload->>'end', d.payload->>'quantityField'
  from public.bom_documents d
  where d.kind in ('STOCK_IN','STOCK_OUT','EXTRA_IN')
    and exists (select 1 from public.bom_members m where m.user_id=auth.uid())
$$;
revoke all on function public.bom_import_coverage() from public, anon;
grant execute on function public.bom_import_coverage() to authenticated;
commit;

