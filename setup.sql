-- ============================================================
-- AMBRAC Agenda de Visitas — Supabase setup
-- Execute no SQL Editor do projeto Supabase
-- ============================================================

CREATE TABLE IF NOT EXISTS agendamentos (
  id                      uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  data_visita             date NOT NULL,
  ordem                   integer,
  categoria               text,
  classificacao           integer,
  cnpj                    text,
  grupo                   text,
  razao_social            text,
  unidade                 text,
  servicos_contrato       text,
  data_documentacao       text,
  visitas_feitas          text,
  psicossocial            text,
  status_documentacao     text,
  autor                   text,
  grazi                   boolean DEFAULT false,
  responsavel_documentacao text,
  tipo_contrato           text,
  data_contrato           text,
  planilha_empresa        text,
  comercial_responsavel   text,
  ordem_servico_unisyst   text,
  visita                  text,
  cidade                  text,
  endereco                text,
  email                   text,
  telefone                text,
  uber                    text,
  data_pagamento          text,
  criado_em               timestamptz DEFAULT now(),
  atualizado_em           timestamptz DEFAULT now(),
  atualizado_por          text
);

-- Índices para performance nas queries mais comuns
CREATE INDEX IF NOT EXISTS idx_agend_data_visita ON agendamentos (data_visita);
CREATE INDEX IF NOT EXISTS idx_agend_status      ON agendamentos (status_documentacao);
CREATE INDEX IF NOT EXISTS idx_agend_autor       ON agendamentos (autor);

-- RLS: permite acesso total para anon (sem login)
ALTER TABLE agendamentos ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_all" ON agendamentos;
CREATE POLICY "anon_all" ON agendamentos
  FOR ALL TO anon
  USING (true)
  WITH CHECK (true);