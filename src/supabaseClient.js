import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = 'https://zfmawuetovlokadzatmw.supabase.co';
const SUPABASE_KEY = 'sb_publishable_z-rOkW5t4A-MY_c7hvRFtg_tWo_60sh';

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
