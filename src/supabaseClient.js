import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://nwdaryuauhziabxlseza.supabase.co'

// Paste your freshly copied eyJ... key inside the quotes
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im53ZGFyeXVhdWh6aWFieGxzZXphIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQ0MjkyODEsImV4cCI6MjEwMDAwNTI4MX0.3o2Slr5mOcGAionCrXxWgTjpvPo-7dYTHnSBD0vyE4Q'

export const supabase = createClient(supabaseUrl, supabaseAnonKey)