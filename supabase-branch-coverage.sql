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
