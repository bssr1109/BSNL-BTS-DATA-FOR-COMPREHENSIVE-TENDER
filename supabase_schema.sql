-- Run this in Supabase SQL Editor after deleting old wrong table
create table if not exists public.sites (
  id bigserial primary key,
  si_no integer,
  ba_name text,
  oa_name text,
  sdca_name text,
  bts_type text,
  tower_type text,
  bts_name text,
  rp_id text unique,
  site_type text,
  sites_having_acs text,
  acs_count integer,
  make text,
  ac_capacity text,
  status_of_ac text,
  ac_pm_cm text,
  free_cooling_available text,
  site_load_gt_10kw text,
  remarks text,
  sde_name text,
  last_updated_by text,
  last_updated_at timestamptz,
  is_completed boolean default false
);

create table if not exists public.dropdown_options (
  field_name text not null,
  option_value text not null,
  sort_order integer default 0,
  primary key (field_name, option_value)
);

insert into public.dropdown_options(field_name, option_value, sort_order) values
('tower_type','GBT',1),('tower_type','RTT',2),('tower_type','RTP',3),('tower_type','GBM',4),
('site_type','Building',1),('site_type','Shelter',2),('site_type','Cage',3),('site_type','Open',4),
('sites_having_acs','Yes',1),('sites_having_acs','No',2),
('status_of_ac','Working',1),('status_of_ac','Faulty',2),('status_of_ac','Not Available',3),
('ac_pm_cm','Required',1),('ac_pm_cm','Not Required',2),
('free_cooling_available','Yes',1),('free_cooling_available','No',2),
('site_load_gt_10kw','Yes',1),('site_load_gt_10kw','No',2)
on conflict (field_name, option_value) do nothing;

alter table public.sites enable row level security;
alter table public.dropdown_options enable row level security;

-- For simple Streamlit app using anon key/service role in secrets.
-- If policies block the app, add policy below or use SERVICE_ROLE key in Streamlit secrets.
do $$ begin
  create policy "allow_read_dropdowns" on public.dropdown_options for select using (true);
exception when duplicate_object then null; end $$;

do $$ begin
  create policy "allow_all_sites_for_app" on public.sites for all using (true) with check (true);
exception when duplicate_object then null; end $$;