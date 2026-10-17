-- Allow branch members to edit shared BOM; preserve source import restrictions.
begin;
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

