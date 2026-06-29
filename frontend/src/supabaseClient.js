import { createClient } from '@supabase/supabase-js'

// Replace these with your actual Supabase project credentials.
// Recommended: store them in a .env file as VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY
const SUPABASE_URL  = import.meta.env.VITE_SUPABASE_URL  || 'https://wvmxlnwfjtesbppojstu.supabase.co'
const SUPABASE_ANON = import.meta.env.VITE_SUPABASE_ANON_KEY || 'sb_publishable_B9De51Ete2xOSG4YlVzgAg_8WbiSI6j'

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON)
