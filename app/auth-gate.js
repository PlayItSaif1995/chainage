/**
 * CHAINAGE — AUTH GATE
 * 
 * This file goes in your /app/ folder alongside index.html (your main programme).
 * It checks if the user is logged in and has active access before showing the app.
 * 
 * How it works:
 * 1. User visits chainage.net/app/index.html
 * 2. This script runs first (loaded at the top of app/index.html)
 * 3. If not logged in → redirected to /login.html
 * 4. If logged in but no active subscription → redirected to /subscribe.html
 * 5. If logged in and has access → app loads normally
 */

// ── PASTE YOUR SUPABASE VALUES HERE ──
const SUPABASE_URL     = 'https://dutwdbjpniiqtifrcvfp.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_7tkCapKdEyLyLcOBDUqQBg_j08PlZb8';

(async function() {
  // Load Supabase if not already loaded
  if (!window.supabase) {
    await loadScript('https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2');
  }

  const { createClient } = window.supabase;
  const sb = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

  // Get current session
  const { data: { session } } = await sb.auth.getSession();

  if (!session) {
    // Not logged in — send to login
    window.location.replace('../login.html');
    return;
  }

  // Check access (beta flag OR active Stripe subscription)
  const { data: access } = await sb
    .from('user_access')
    .select('is_beta, is_active, plan')
    .eq('user_id', session.user.id)
    .single();

  if (!access || (!access.is_beta && !access.is_active)) {
    // No active access — send to subscribe page
    window.location.replace('../subscribe.html');
    return;
  }

  // ✓ User is authenticated and has access — app will load
  // Expose user info globally for use in the app if needed
  window.chainageUser = {
    id:    session.user.id,
    email: session.user.email,
    name:  session.user.user_metadata?.first_name || '',
    plan:  access.plan || 'individual',
    isBeta: access.is_beta || false,
  };

  // Add sign-out capability
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
