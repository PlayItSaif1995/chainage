// ── PASTE YOUR SUPABASE VALUES HERE ──
const SUPABASE_URL      = 'https://dutwdbjpniiqtifrcvfp.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_7tkCapKdEyLyLcOBDUqQBg_j08PlZb8';

(async function() {
  if (!window.supabase) {
    await loadScript('https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2');
  }

  const { createClient } = window.supabase;
  const sb = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

  const { data: { session } } = await sb.auth.getSession();

  if (!session) {
    window.location.replace('../login.html');
    return;
  }

  // Use limit(1) instead of single() to handle duplicate rows without error
  const { data: rows } = await sb
    .from('user_access')
    .select('is_beta, is_active, plan')
    .eq('user_id', session.user.id)
    .limit(1);

  const access = rows && rows.length > 0 ? rows[0] : null;

  if (!access || (!access.is_beta && !access.is_active)) {
    window.location.replace('../subscribe.html');
    return;
  }

  window.chainageUser = {
    id:     session.user.id,
    email:  session.user.email,
    name:   session.user.user_metadata?.first_name || '',
    plan:   access.plan || 'individual',
    isBeta: access.is_beta || false,
  };

  window.chainageSignOut = async () => {
    await sb.auth.signOut();
    window.location.replace('../login.html');
  };

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src; s.onload = resolve; s.onerror = reject;
      document.head.appendChild(s);
    });
  }
})();
