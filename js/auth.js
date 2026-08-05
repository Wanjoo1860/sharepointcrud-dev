// ============================================================
// 인증 / 로그인 / 토큰 관리
// ============================================================

// ---- 앱 초기화 (즉시 실행) ----
(async function initApp() {
    // 1. Teams 환경 감지
    try {
        await Promise.race([
            microsoftTeams.app.initialize(),
            new Promise(function (_, reject) {
                setTimeout(function () { reject(new Error('Not in Teams')); }, 500);
            })
        ]);
        APP.isInTeams = true;
        document.getElementById('envBadge').innerText = 'Teams';
        document.getElementById('envBadge').className = 'env-badge env-teams';
        console.log('[ENV] Teams 환경 감지됨');
    } catch (e) {
        APP.isInTeams = false;
        document.getElementById('envBadge').innerText = 'Browser';
        document.getElementById('envBadge').className = 'env-badge env-browser';
        console.log('[ENV] 일반 브라우저 환경');
    }

    // 2. MSAL 초기화
    if (APP.isInTeams) {
        try {
            APP.msalInstance = await msal.createNestablePublicClientApplication({
                auth: {
                    clientId: CONFIG.clientId,
                    authority: 'https://login.microsoftonline.com/' + CONFIG.tenantId,
                    supportsNestedAppAuth: true
                }
            });
            APP.msalReady = true;
            console.log('[MSAL] NAA 초기화 성공');
        } catch (e) {
            console.error('[MSAL] NAA 초기화 실패:', e.message);
            showStatus('NAA 초기화 실패: ' + e.message, 'error');
            return;
        }
    } else {
        APP.msalInstance = new msal.PublicClientApplication({
            auth: {
                clientId: CONFIG.clientId,
                authority: 'https://login.microsoftonline.com/' + CONFIG.tenantId,
                redirectUri: window.location.origin + window.location.pathname
            },
            cache: { cacheLocation: 'localStorage' }
        });

        await APP.msalInstance.initialize();
        console.log('[MSAL] initialize() 완료');

        try {
            var response = await APP.msalInstance.handleRedirectPromise();
            if (response && response.account) {
                console.log('[MSAL] 리디렉트 복귀 - 토큰 수신');
                APP.msalInstance.setActiveAccount(response.account);
                if (response.accessToken) {
                    APP.accessToken = response.accessToken;
                } else {
                    await getToken();
                }
                APP.msalReady = true;
                onLoginSuccess();
                return;
            }
        } catch (e) {
            console.error('[MSAL] 리디렉트 처리 오류:', e.message);
        }

        APP.msalReady = true;
        console.log('[MSAL] 브라우저 MSAL 준비 완료');
    }

    // 3. 기존 세션 자동 로그인
    if (APP.isInTeams) {
        showStatus('자동 로그인 중...', 'info');
        try {
            var tokenResponse = await APP.msalInstance.acquireTokenSilent({ scopes: CONFIG.scopes });
            APP.accessToken = tokenResponse.accessToken;
            if (tokenResponse.account) APP.msalInstance.setActiveAccount(tokenResponse.account);
            onLoginSuccess();
        } catch (silentError) {
            try {
                var tokenResponse = await APP.msalInstance.acquireTokenPopup({ scopes: CONFIG.scopes });
                APP.accessToken = tokenResponse.accessToken;
                APP.msalInstance.setActiveAccount(tokenResponse.account);
                onLoginSuccess();
            } catch (popupError) {
                showStatus('자동 로그인 실패. 로그인 버튼을 클릭하세요.', 'error');
            }
        }
    } else {
        var accounts = APP.msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            APP.msalInstance.setActiveAccount(accounts[0]);
            try {
                await getToken();
                onLoginSuccess();
            } catch (e) {
                console.log('[MSAL] 자동 로그인 실패:', e.message);
            }
        }
    }
})();

// ---- 수동 로그인 ----
async function login() {
    if (!APP.msalReady) {
        showStatus('MSAL 초기화 중... 잠시 후 다시 시도하세요.', 'error');
        return;
    }

    try {
        if (APP.isInTeams) {
            try {
                var response = await APP.msalInstance.acquireTokenSilent({ scopes: CONFIG.scopes });
                APP.accessToken = response.accessToken;
                APP.msalInstance.setActiveAccount(response.account);
            } catch (e) {
                var response = await APP.msalInstance.acquireTokenPopup({ scopes: CONFIG.scopes });
                APP.accessToken = response.accessToken;
                APP.msalInstance.setActiveAccount(response.account);
            }
        } else {
            try {
                var response = await APP.msalInstance.loginPopup({ scopes: CONFIG.scopes });
                APP.msalInstance.setActiveAccount(response.account);
                if (response.accessToken) {
                    APP.accessToken = response.accessToken;
                } else {
                    await getToken();
                }
            } catch (popupError) {
                console.warn('[MSAL] Popup 실패, redirect 폴백:', popupError.message);
                await APP.msalInstance.loginRedirect({ scopes: CONFIG.scopes });
                return;
            }
        }
        onLoginSuccess();
    } catch (e) {
        showStatus('로그인 실패: ' + e.message, 'error');
    }
}

// ---- 토큰 획득 ----
async function getToken() {
    var account = APP.msalInstance.getActiveAccount();
    if (!account) {
        var accounts = APP.msalInstance.getAllAccounts();
        if (accounts.length === 0) throw new Error('로그인 필요');
        account = accounts[0];
        APP.msalInstance.setActiveAccount(account);
    }

    try {
        var response = await APP.msalInstance.acquireTokenSilent({
            scopes: CONFIG.scopes,
            account: account
        });
        APP.accessToken = response.accessToken;
    } catch (e) {
        console.warn('[MSAL] Silent 토큰 실패:', e.message);
        if (APP.isInTeams) {
            var response = await APP.msalInstance.acquireTokenPopup({ scopes: CONFIG.scopes });
            APP.accessToken = response.accessToken;
        } else {
            try {
                var response = await APP.msalInstance.acquireTokenPopup({ scopes: CONFIG.scopes });
                APP.accessToken = response.accessToken;
            } catch (popupErr) {
                await APP.msalInstance.acquireTokenRedirect({ scopes: CONFIG.scopes });
            }
        }
    }
}

// ---- 로그아웃 ----
function logout() {
    if (APP.isInTeams) {
        APP.msalInstance.setActiveAccount(null);
        APP.accessToken = null;
        APP.currentUserRole = 'user';
        document.getElementById('btnLogin').style.display = '';
        document.getElementById('btnLogout').style.display = 'none';
        document.getElementById('loginStatus').innerText = '로그인이 필요합니다.';
        document.getElementById('btnRefresh').disabled = true;
        document.getElementById('btnSave').disabled = true;
        document.getElementById('adminPanel').style.display = 'none';
        document.getElementById('tableBody').innerHTML = '<tr><td colspan="5" style="text-align:center;padding:30px;color:#999;">로그인 후 데이터를 불러옵니다.</td></tr>';
        showStatus('로그아웃 완료', 'info');
    } else {
        var account = APP.msalInstance.getActiveAccount();
        APP.msalInstance.logoutPopup({
            account: account,
            postLogoutRedirectUri: window.location.origin + window.location.pathname
        }).catch(function () {
            APP.msalInstance.logoutRedirect({
                postLogoutRedirectUri: window.location.origin + window.location.pathname
            });
        });
    }
}

// ---- 로그인 성공 처리 ----
function onLoginSuccess() {
    document.getElementById('btnLogin').style.display = 'none';
    document.getElementById('btnLogout').style.display = '';
    var account = APP.msalInstance.getActiveAccount();
    APP.currentUserEmail = account ? account.username : '';
    document.getElementById('loginStatus').innerText = '✓ ' + APP.currentUserEmail;
    document.getElementById('btnRefresh').disabled = false;
    document.getElementById('btnSave').disabled = false;
    showStatus('로그인 성공', 'success');

    // 역할 확인 후 데이터 로드
    checkUserRole().then(function () {
        getSiteAndList();
    });
}
