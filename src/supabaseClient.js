import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://nwdaryuauhziabxtlseza.supabase.co'
const supabaseAnonKey = 'sb_publishable_qu0PTxwHp2hrAPoZimohXA_ajNkPfH6'

export const supabase = createClient(supabaseUrl, supabaseAnonKey)