# Cloud Sync Setup (Supabase)

## 1) Create Supabase project
- Go to https://supabase.com
- Create a new project

## 2) Create table in SQL Editor
Run this SQL:

```sql
create table if not exists public.app_state (
  app_id text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now(),
  updated_by text
);

alter table public.app_state enable row level security;

create policy "allow read app_state"
on public.app_state
for select
using (true);

create policy "allow insert app_state"
on public.app_state
for insert
with check (true);

create policy "allow update app_state"
on public.app_state
for update
using (true)
with check (true);
```

## 3) Get project URL and anon key
- Supabase dashboard -> Project Settings -> API
- Copy:
  - Project URL
  - anon public key

## 4) Configure in app
- Open app -> Admin Settings
- Fill:
  - Supabase Project URL
  - Supabase Anon Key
  - Cloud App ID (same value on all devices)
- Click `Save Cloud Config`

## 5) Sync usage
- `Push To Cloud`: upload current device data to cloud
- `Pull From Cloud`: download latest cloud data to this device
- `Enable Cloud Auto Sync`: pushes local changes + checks cloud every minute

## Notes
- Use the same `Cloud App ID` on all devices that should share one dataset.
- If two devices edit at the same time, latest update time wins.
- Keep backup downloads enabled as safety.
