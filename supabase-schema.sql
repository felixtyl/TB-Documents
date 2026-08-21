-- ============================================================
-- Production Document Center — Supabase schema
-- Run this once in Supabase Dashboard > SQL Editor > New Query
-- ============================================================

-- ---------- Profiles (one row per signed-up person) ----------
create table public.profiles (
  id uuid references auth.users on delete cascade primary key,
  name text not null,
  role text not null default 'member' check (role in ('admin','member')),
  can_build boolean not null default false,       -- template builder access
  can_build_docs boolean not null default false,  -- document builder access
  active boolean not null default true,           -- deactivated users are locked out
  created_at timestamptz not null default now()
);

alter table public.profiles enable row level security;

-- Helper: check if the current user is an admin, without recursive RLS issues
create or replace function public.is_admin()
returns boolean
language sql
security definer
set search_path = public
as $$
  select exists (
    select 1 from public.profiles where id = auth.uid() and role = 'admin'
  );
$$;

create policy "profiles readable by any signed-in user"
on public.profiles for select
to authenticated
using (true);

create policy "only admins can change profiles"
on public.profiles for update
to authenticated
using (public.is_admin())
with check (public.is_admin());

-- Auto-create a profile row whenever someone signs up.
-- The very first person to sign up becomes Admin automatically.
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  is_first boolean;
begin
  select not exists(select 1 from public.profiles) into is_first;
  insert into public.profiles (id, name, role, can_build, can_build_docs, active)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'name', split_part(new.email, '@', 1)),
    case when is_first then 'admin' else 'member' end,
    is_first,
    is_first,
    true
  );
  return new;
end;
$$;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- ---------- Documents ----------
create table public.documents (
  id uuid primary key default gen_random_uuid(),
  title text not null,
  type text not null default 'SOP',
  department text,
  owner text,
  revision text default 'A',
  status text not null default 'Draft' check (status in ('Draft','Pending Approval','Approved')),
  effective_date date,
  expiry_date date,
  content text,
  attachments jsonb not null default '[]'::jsonb,
  created_by uuid references auth.users,
  created_by_name text,
  created_at timestamptz not null default now(),
  updated_by_name text,
  updated_at timestamptz not null default now()
);
alter table public.documents enable row level security;

create policy "signed-in users can read documents"
on public.documents for select to authenticated using (true);

create policy "doc builders can insert documents"
on public.documents for insert to authenticated
with check (exists (
  select 1 from public.profiles p
  where p.id = auth.uid() and p.active and (p.role = 'admin' or p.can_build_docs)
));

create policy "doc builders can update documents"
on public.documents for update to authenticated
using (exists (
  select 1 from public.profiles p
  where p.id = auth.uid() and p.active and (p.role = 'admin' or p.can_build_docs)
));

create policy "admins can delete documents"
on public.documents for delete to authenticated
using (public.is_admin());

-- ---------- Templates ----------
create table public.templates (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  category text not null default 'Inspection Checklist',
  department text,
  fields jsonb not null default '[]'::jsonb,
  created_by_name text,
  created_at timestamptz not null default now(),
  updated_by_name text,
  updated_at timestamptz not null default now()
);
alter table public.templates enable row level security;

create policy "signed-in users can read templates"
on public.templates for select to authenticated using (true);

create policy "template builders can insert templates"
on public.templates for insert to authenticated
with check (exists (
  select 1 from public.profiles p
  where p.id = auth.uid() and p.active and (p.role = 'admin' or p.can_build)
));

create policy "template builders can update templates"
on public.templates for update to authenticated
using (exists (
  select 1 from public.profiles p
  where p.id = auth.uid() and p.active and (p.role = 'admin' or p.can_build)
));

create policy "admins can delete templates"
on public.templates for delete to authenticated
using (public.is_admin());

-- ---------- Submissions (filled-out templates) ----------
create table public.submissions (
  id uuid primary key default gen_random_uuid(),
  template_id uuid references public.templates(id) on delete set null,
  template_name text,
  category text,
  department text,
  fields_snapshot jsonb not null default '[]'::jsonb,
  values jsonb not null default '{}'::jsonb,
  filled_by_name text,
  filled_at date not null default current_date,
  filled_at_time timestamptz not null default now()
);
alter table public.submissions enable row level security;

create policy "signed-in users can read submissions"
on public.submissions for select to authenticated using (true);

create policy "signed-in users can submit forms"
on public.submissions for insert to authenticated with check (true);

-- ---------- Storage bucket for attachments ----------
insert into storage.buckets (id, name, public)
values ('attachments', 'attachments', false)
on conflict (id) do nothing;

create policy "signed-in users can read attachments"
on storage.objects for select to authenticated
using (bucket_id = 'attachments');

create policy "signed-in users can upload attachments"
on storage.objects for insert to authenticated
with check (bucket_id = 'attachments');

create policy "signed-in users can delete attachments"
on storage.objects for delete to authenticated
using (bucket_id = 'attachments');
