import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://nwdaryuauhziabxtlseza.supabase.co'
const supabaseAnonKey =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im53ZGFyeXVhdWh6aWFieHRsc2V6YSIsInJvbGUiOiJhb24iLCJpYXQiOjE3MzODQ0MjkyODEsImV4cCI6MjEwMDAwNTI4MX0.3o2S1r5mOcGAIonCrXwgTjpvPo-7dYTHnSBD0vyE4Q'

export const supabase = createClient(supabaseUrl, supabaseAnonKey)