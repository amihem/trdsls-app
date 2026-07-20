import { createClient } from '@supabase/supabase-js'

// Environment variable padhega, agar empty milega toh direct URL aur Key use karega
const supabaseUrl = process.env.REACT_APP_SUPABASE_URL || 'https://nwdaryuauhziabxlseza.supabase.co'
const supabaseAnonKey = process.env.REACT_APP_SUPABASE_ANON_KEY || 'sb_publishable_qu0PTxwHp2hrAPoZimohXA_ajNkPfH6'

export const supabase = createClient(supabaseUrl, supabaseAnonKey)