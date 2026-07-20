import { createClient } from '@supabase/supabase-js'

// Correct URL with 'l' (nwdaryuauhziabxlseza)
const supabaseUrl = 'https://nwdaryuauhziabxlseza.supabase.co'

const supabaseAnonKey =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im53ZGFyeXVhdWh6aWFieGxzZXphIiwicm9sZSI6ImFub24iLCJpYXQiOjE3MzODQ0MjkyODEsImV4cCI6MjEwMDAwNTI4MX0.3o2S1r5mOcGAIonCrXwgTjpvPo-7dYTHnSBD0vyE4Q'

export const supabase = createClient(supabaseUrl, supabaseAnonKey)