import { createClient } from '@supabase/supabase-js'

// Replace these with your actual Supabase project credentials.
// Recommended: store them in a .env file as VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY
const SUPABASE_URL  = import.meta.env.VITE_SUPABASE_URL  || 'https://YOUR_PROJECT.supabase.co'
const SUPABASE_ANON = import.meta.env.VITE_SUPABASE_ANON_KEY || 'YOUR_ANON_KEY'

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON)
