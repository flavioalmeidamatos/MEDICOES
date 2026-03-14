# AGENTS.md

## Projeto
cadclientes

## Objetivo
Aplicação de cadastro e gestão de clientes.

## Stack esperada
Verificar no projeto:
- React
- TypeScript
- Vite
- Supabase
- Vercel

## Regras
- Antes de editar, mapear impacto em autenticação, rotas, formulário, listagem e Supabase.
- Preferir alterações pequenas.
- Não alterar `.env.local` automaticamente.
- Toda mudança de banco deve virar migration versionada.
- Não rodar deploy em produção sem confirmação.
- Não misturar refatoração ampla com correção funcional.

## Áreas normalmente críticas
- `App.tsx`
- `contexts/AuthContext.tsx`
- `supabase.ts` ou `src/lib/supabase.ts`
- `pages/ClientList.tsx`
- `pages/ClientForm.tsx`
- `components/Layout.tsx`
- `types.ts`
- `constants.ts`

## Build
```bash
npm run build