// ============================================================
// ★★★ 설정 값 ★★★
// ============================================================
var CLIENT_ID = '4aeb92a9-8ef6-476a-a419-9125032309fd';
var TENANT_ID = 'cc13b6f1-ef21-479f-9853-3e2dffa71d6b';
var SITE_HOSTNAME = 'globalsoft.sharepoint.com';
var SITE_PATH = '/sites/Dev896';
var LIST_NAME = 'TestData';
var SCOPES = ['User.Read', 'Sites.ReadWrite.All'];
// ============================================================

var GRAPH_URL = 'https://graph.microsoft.com/v1.0';
var DATA = [];
var siteId = null;
var listId = null;
var accessToken = null;
var msalInstance = null;
var isInTeams = false;
var msalReady = false;

// ---- 앱 초기화 ----
(async function () {
    // 1. Teams 환경 감지 (0.5초 타임아웃 적용)
    try {
        await Promise.race([
            microsoftTeams.app.initialize(),
            new Promise(function (_, reject) {
                setTimeout(function () { reject(new Error('Not in Teams')); }, 500);
            })
        ]);
        isInTeams = true;
        document.getElementById('envBadge').innerText = 'Teams';
        document.getElementById('envBadge').className = 'env-badge env-teams';
        console.log('[ENV] Teams 환경 감지됨 - NAA 모드');
    } catch (e) {
        isInTeams = false;
        document.getElementById('envBadge').innerText = 'Browser';
        document.getElementById('envBadge').className = 'env-badge env-browser';
        console.log('[ENV] 일반 브라우저 환경');
    }

    // 2. MSAL 초기화
    if (isInTeams) {
        try {
            msalInstance = await msal.createNestablePublicClientApplication({
                auth: {
                    clientId: CLIENT_ID,
                    authority: 'https://login.microsoftonline.com/' + TENANT_ID,
                    supportsNestedAppAuth: true
                }
            });
            msalReady = true;
            console.log('[MSAL] NAA 초기화 성공');
        } catch (e) {
            console.error('[MSAL] NAA 초기화 실패:', e.message);
            showStatus('NAA 초기화 실패: ' + e.message, 'error');
            return;
        }
    } else {
        msalInstance = new msal.PublicClientApplication({
            auth: {
                clientId: CLIENT_ID,
                authority: 'https://login.microsoftonline.com/' + TENANT_ID,
                redirectUri: window.location.origin + window.location.pathname
            },
            cache: {
                cacheLocation: 'localStorage'
            }
        });

        await msalInstance.initialize();
        console.log('[MSAL] initialize() 완료');

        try {
            var response = await msalInstance.handleRedirectPromise();
            if (response && response.account) {
                console.log('[MSAL] 리디렉트 복귀 - 토큰 수신');
                msalInstance.setActiveAccount(response.account);
                if (response.accessToken) {
                    accessToken = response.accessToken;
                } else {
                    await getToken();
                }
                msalReady = true;
                onLoginSuccess();
                return;
            }
        } catch (e) {
            console.error('[MSAL] 리디렉트 처리 오류:', e.message);
        }

        msalReady = true;
        console.log('[MSAL] 브라우저 MSAL 준비 완료');
    }

    // 3. 기존 세션이 있으면 자동 로그인 시도
    if (isInTeams) {
        showStatus('자동 로그인 중...', 'info');
        try {
            var tokenResponse = await msalInstance.acquireTokenSilent({
                scopes: SCOPES
            });
            accessToken = tokenResponse.accessToken;
            if (tokenResponse.account) {
                msalInstance.setActiveAccount(tokenResponse.account);
            }
            onLoginSuccess();
        } catch (silentError) {
            console.log('[MSAL] Teams silent 실패, popup 시도:', silentError.message);
            try {
                var tokenResponse = await msalInstance.acquireTokenPopup({
                    scopes: SCOPES
                });
                accessToken = tokenResponse.accessToken;
                msalInstance.setActiveAccount(tokenResponse.account);
                onLoginSuccess();
            } catch (popupError) {
                showStatus('자동 로그인 실패. 로그인 버튼을 클릭하세요.', 'error');
                console.error('[MSAL] Teams 자동 로그인 실패:', popupError.message);
            }
        }
    } else {
        var accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            msalInstance.setActiveAccount(accounts[0]);
            try {
                await getToken();
                onLoginSuccess();
            } catch (e) {
                console.log('[MSAL] 자동 로그인 실패, 수동 로그인 필요:', e.message);
            }
        }
    }
})();

// ---- 로그인 (수동) ----
async function login() {
    if (!msalReady) {
        showStatus('MSAL 초기화 중... 잠시 후 다시 시도하세요.', 'error');
        return;
    }

    try {
        if (isInTeams) {
            try {
                var response = await msalInstance.acquireTokenSilent({ scopes: SCOPES });
                accessToken = response.accessToken;
                msalInstance.setActiveAccount(response.account);
            } catch (e) {
                var response = await msalInstance.acquireTokenPopup({ scopes: SCOPES });
                accessToken = response.accessToken;
                msalInstance.setActiveAccount(response.account);
            }
        } else {
            try {
                var response = await msalInstance.loginPopup({ scopes: SCOPES });
                msalInstance.setActiveAccount(response.account);
                if (response.accessToken) {
                    accessToken = response.accessToken;
                } else {
                    await getToken();
                }
            } catch (popupError) {
                console.warn('[MSAL] Popup 실패, redirect로 폴백:', popupError.message);
                await msalInstance.loginRedirect({ scopes: SCOPES });
                return;
            }
        }

        onLoginSuccess();
    } catch (e) {
        showStatus('로그인 실패: ' + e.message, 'error');
        console.error('[MSAL] 로그인 실패:', e);
    }
}

// ---- 토큰 획득 ----
async function getToken() {
    var account = msalInstance.getActiveAccount();
    if (!account) {
        var accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) throw new Error('로그인 필요');
        account = accounts[0];
        msalInstance.setActiveAccount(account);
    }

    try {
        var response = await msalInstance.acquireTokenSilent({
            scopes: ['Sites.ReadWrite.All'],
            account: account
        });
        accessToken = response.accessToken;
    } catch (e) {
        console.warn('[MSAL] Silent 토큰 실패, interactive 시도:', e.message);
        if (isInTeams) {
            var response = await msalInstance.acquireTokenPopup({
                scopes: ['Sites.ReadWrite.All']
            });
            accessToken = response.accessToken;
        } else {
            try {
                var response = await msalInstance.acquireTokenPopup({
                    scopes: ['Sites.ReadWrite.All']
                });
                accessToken = response.accessToken;
            } catch (popupErr) {
                console.warn('[MSAL] Token popup 실패, redirect:', popupErr.message);
                await msalInstance.acquireTokenRedirect({
                    scopes: ['Sites.ReadWrite.All']
                });
            }
        }
    }
}

// ---- 로그아웃 ----
function logout() {
    if (isInTeams) {
        msalInstance.setActiveAccount(null);
        accessToken = null;
        document.getElementById('btnLogin').style.display = '';
        document.getElementById('btnLogout').style.display = 'none';
        document.getElementById('loginStatus').innerText = '로그인이 필요합니다.';
        document.getElementById('btnRefresh').disabled = true;
        document.getElementById('btnSave').disabled = true;
        document.getElementById('tableBody').innerHTML = '<tr><td colspan="5" style="text-align:center;padding:30px;color:#999;">로그인 후 데이터를 불러옵니다.</td></tr>';
        showStatus('로그아웃 완료', 'info');
    } else {
        var account = msalInstance.getActiveAccount();
        msalInstance.logoutPopup({
            account: account,
            postLogoutRedirectUri: window.location.origin + window.location.pathname
        }).catch(function () {
            msalInstance.logoutRedirect({
                postLogoutRedirectUri: window.location.origin + window.location.pathname
            });
        });
    }
}

// ---- 로그인 성공 처리 ----
function onLoginSuccess() {
    document.getElementById('btnLogin').style.display = 'none';
    document.getElementById('btnLogout').style.display = '';
    var account = msalInstance.getActiveAccount();
    document.getElementById('loginStatus').innerText = '✓ ' + (account ? account.username : '로그인됨');
    document.getElementById('btnRefresh').disabled = false;
    document.getElementById('btnSave').disabled = false;
    showStatus('로그인 성공: ' + (account ? account.username : ''), 'success');
    getSiteAndList();
}

// ---- Graph API 호출 헬퍼 ----
async function graphCall(url, method, body) {
    await getToken();
    var options = {
        method: method || 'GET',
        headers: {
            'Authorization': 'Bearer ' + accessToken,
            'Content-Type': 'application/json'
        }
    };
    if (body) options.body = JSON.stringify(body);
    var r = await fetch(url, options);
    if (method === 'DELETE' && r.status === 204) return null;
    if (!r.ok) {
        var err = await r.text();
        throw new Error('HTTP ' + r.status + ': ' + err);
    }
    if (r.status === 204) return null;
    return await r.json();
}

// ---- Site ID, List ID 조회 ----
async function getSiteAndList() {
    try {
        showStatus('사이트 정보 조회 중...', 'info');
        var site = await graphCall(GRAPH_URL + '/sites/' + SITE_HOSTNAME + ':' + SITE_PATH);
        siteId = site.id;
        var list = await graphCall(GRAPH_URL + '/sites/' + siteId + '/lists/' + LIST_NAME);
        listId = list.id;
        showStatus('연결 완료: ' + LIST_NAME, 'success');
        getItems();
    } catch (e) {
        showStatus('사이트/목록 조회 실패: ' + e.message, 'error');
    }
}

// ---- READ ----
async function getItems() {
    try {
        showStatus('데이터 불러오는 중...', 'info');
        var result = await graphCall(
            GRAPH_URL + '/sites/' + siteId + '/lists/' + listId + '/items?$expand=fields&$top=500'
        );
        DATA = result.value;
        renderTable();
        showStatus('로드 완료 (' + DATA.length + '건)', 'success');
        document.getElementById('countInfo').innerText = '총 ' + DATA.length + '건';
    } catch (e) {
        showStatus('로드 실패: ' + e.message, 'error');
    }
}

// ---- CREATE ----
async function createItem(name, dept, pos) {
    var body = { fields: { Title: name } };
    body.fields['_xbd80__xc11c_'] = dept;
    body.fields['_xc9c1__xae09_'] = pos;
    await graphCall(
        GRAPH_URL + '/sites/' + siteId + '/lists/' + listId + '/items',
        'POST', body
    );
}

// ---- UPDATE ----
async function updateItem(id, name, dept, pos) {
    var body = { Title: name };
    body['_xbd80__xc11c_'] = dept;
    body['_xc9c1__xae09_'] = pos;
    await graphCall(
        GRAPH_URL + '/sites/' + siteId + '/lists/' + listId + '/items/' + id + '/fields',
        'PATCH', body
    );
}

// ---- DELETE ----
async function deleteItemFromList(id) {
    await graphCall(
        GRAPH_URL + '/sites/' + siteId + '/lists/' + listId + '/items/' + id,
        'DELETE'
    );
}

// ---- SAVE (Add or Edit) ----
async function saveItem() {
    var editId = document.getElementById('editId').value;
    var name = document.getElementById('inputName').value.trim();
    var dept = document.getElementById('inputDept').value.trim();
    var pos = document.getElementById('inputPosition').value.trim();
    if (!name || !dept || !pos) {
        showStatus('모든 항목을 입력해주세요.', 'error');
        return;
    }
    try {
        if (editId) {
            await updateItem(editId, name, dept, pos);
            showStatus('ID ' + editId + ' 수정 완료', 'success');
        } else {
            await createItem(name, dept, pos);
            showStatus('새 항목 추가 완료', 'success');
        }
        cancelEdit();
        await getItems();
    } catch (e) {
        showStatus('저장 실패: ' + e.message, 'error');
    }
}

// ---- DELETE (UI) ----
async function deleteItem(id) {
    if (!confirm('ID ' + id + ' 항목을 삭제하시겠습니까?')) return;
    try {
        await deleteItemFromList(id);
        showStatus('ID ' + id + ' 삭제 완료', 'success');
        await getItems();
    } catch (e) {
        showStatus('삭제 실패: ' + e.message, 'error');
    }
}

// ---- RENDER ----
function renderTable() {
    var tbody = document.getElementById('tableBody');
    if (DATA.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:30px;color:#999;">데이터가 없습니다.</td></tr>';
        return;
    }
    var html = '';
    DATA.forEach(function (item) {
        var f = item.fields;
        var id = f.id || item.id;
        html += '<tr>';
        html += '<td>' + id + '</td>';
        html += '<td>' + (f.Title || '') + '</td>';
        html += '<td>' + (f['_xbd80__xc11c_'] || '') + '</td>';
        html += '<td>' + (f['_xc9c1__xae09_'] || '') + '</td>';
        html += '<td>';
        html += '<button class="btn btn-warning" onclick="editItem(\'' + id + '\')">수정</button> ';
        html += '<button class="btn btn-danger" onclick="deleteItem(\'' + id + '\')">삭제</button>';
        html += '</td>';
        html += '</tr>';
    });
    tbody.innerHTML = html;
}

// ---- EDIT ----
function editItem(id) {
    var item = DATA.find(function (i) { return (i.fields.id || i.id) == id; });
    if (!item) return;
    var f = item.fields;
    document.getElementById('editId').value = id;
    document.getElementById('inputName').value = f.Title || '';
    document.getElementById('inputDept').value = f['_xbd80__xc11c_'] || '';
    document.getElementById('inputPosition').value = f['_xc9c1__xae09_'] || '';
    document.getElementById('formTitle').innerText = '데이터 수정 (ID: ' + id + ')';
}

// ---- CANCEL ----
function cancelEdit() {
    document.getElementById('editId').value = '';
    document.getElementById('inputName').value = '';
    document.getElementById('inputDept').value = '';
    document.getElementById('inputPosition').value = '';
    document.getElementById('formTitle').innerText = '데이터 추가';
}

// ---- STATUS ----
function showStatus(msg, type) {
    var el = document.getElementById('status');
    el.innerText = msg;
    el.className = 'status show ' + type;
    if (type !== 'info') {
        setTimeout(function () { el.className = 'status'; }, 4000);
    }
}
