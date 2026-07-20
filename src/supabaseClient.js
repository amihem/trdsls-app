import { createClient } from '@supabase/supabase-js'

// Reads from environment variables, or falls back directly to your Supabase credentials
const supabaseUrl = process.env.REACT_APP_SUPABASE_URL || 'https://nwdaryuauhziabxlseza.supabase.co'
const supabaseAnonKey = process.env.REACT_APP_SUPABASE_ANON_KEY || 'sb_publishable_qu0PTxwHp2hrAPoZimohXA_ajNkPfH6'

export const supabase = createClient(supabaseUrl, supabaseAnonKey)