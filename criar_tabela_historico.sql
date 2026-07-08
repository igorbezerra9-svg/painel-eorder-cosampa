-- Rode isso uma vez no SQL Editor do Supabase (painel do projeto -> SQL Editor -> New query -> Run)
-- Cria a tabela que guarda o fechamento diário (últimos 30 dias) de cada região.

create table if not exists snapshots_historico (
  regiao text not null,
  data date not null,
  dados jsonb not null,
  salvo_em timestamptz not null default now(),
  primary key (regiao, data)
);

alter table snapshots_historico enable row level security;

create policy "anon select historico" on snapshots_historico
  for select using (true);
create policy "anon insert historico" on snapshots_historico
  for insert with check (true);
create policy "anon update historico" on snapshots_historico
  for update using (true);
create policy "anon delete historico" on snapshots_historico
  for delete using (true);
