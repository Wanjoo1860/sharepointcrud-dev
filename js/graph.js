// ============================================================
// Graph API 호출 헬퍼
// ============================================================

/**
 * Graph API 범용 호출 함수
 * @param {string} url - API URL
 * @param {string} method - HTTP 메서드 (GET, POST, PATCH, DELETE)
 * @param {object} body - 요청 본문 (선택)
 * @returns {object|null} - 응답 JSON 또는 null
 */
async function graphCall(url, method, body) {
    await getToken();
    var options = {
        method: method || 'GET',
        headers: {
            'Authorization': 'Bearer ' + APP.accessToken,
            'Content-Type': 'application/json'
        }
    };
    if (body) {
        options.body = JSON.stringify(body);
    }

    var response = await fetch(url, options);

    if (method === 'DELETE' && response.status === 204) return null;
    if (!response.ok) {
        var err = await response.text();
        throw new Error('HTTP ' + response.status + ': ' + err);
    }
    if (response.status === 204) return null;
    return await response.json();
}

/**
 * Graph API GET 호출 (토큰 자동 포함)
 * @param {string} url - API URL
 * @returns {object} - 응답 JSON
 */
async function graphGet(url) {
    await getToken();
    var response = await fetch(url, {
        headers: { 'Authorization': 'Bearer ' + APP.accessToken }
    });
    if (!response.ok) {
        var err = await response.text();
        throw new Error('HTTP ' + response.status + ': ' + err);
    }
    return await response.json();
}

/**
 * Graph API POST 호출 (응답 코드만 확인)
 * @param {string} url - API URL
 * @param {object} body - 요청 본문
 * @returns {Response} - fetch Response 객체
 */
async function graphPost(url, body) {
    await getToken();
    return await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': 'Bearer ' + APP.accessToken,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(body)
    });
}

/**
 * Graph API DELETE 호출
 * @param {string} url - API URL
 * @returns {Response} - fetch Response 객체
 */
async function graphDelete(url) {
    await getToken();
    return await fetch(url, {
        method: 'DELETE',
        headers: { 'Authorization': 'Bearer ' + APP.accessToken }
    });
}
