-- ============================================================
-- init.sql (Supabase / Postgres)
-- ============================================================
create extension if not exists "uuid-ossp";

drop table if exists public.hilos;
drop table if exists public.enlaces;

create table public.enlaces (
    id uuid default uuid_generate_v4() primary key,
    nombre_enlace text not null,
    tipo_cable text,
    capacidad int default 24,
    longitud_total text,

    origen_a text,
    sala_a text,
    rack_a text,
    posicion_a text,

    origen_b text,
    sala_b text,
    rack_b text,
    posicion_b text,

    completado_b boolean default false,
    creado_at timestamp with time zone default now()
);

create table public.hilos (
    id uuid default uuid_generate_v4() primary key,
    enlace_id uuid references public.enlaces(id) on delete cascade,
    nro_hilo int not null,
    desc_a text,
    desc_b text
);

create index if not exists hilos_enlace_id_idx on public.hilos(enlace_id);
create unique index if not exists hilos_enlace_nro_unique on public.hilos(enlace_id, nro_hilo);
