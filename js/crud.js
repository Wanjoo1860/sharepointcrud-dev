// ============================================================
// SharePoint 목록 CRUD (Create, Read, Update, Delete)
// ============================================================

/**
 * Site ID와 List ID 조회
 */
async function getSiteAndList() {
    try {
        showStatus('사이트 정보 조회 중...', 'info');
        var site = await graphCall(
            CONFIG.graphUrl + '/sites/' + CONFIG.siteHostname + ':' + CONFIG.sitePath
        );
        APP.siteId = site.id;

        var list = await graphCall(
            CONFIG.graphUrl + '/sites/' + APP.siteId + '/lists/' + CONFIG.listName
        );
        APP.listId = list.id;

        showStatus('연결 완료: ' + CONFIG.listName, 'success');
        getItems();
    } catch (e) {
        showStatus('사이트/목록 조회 실패: ' + e.message, 'error');
    }
}

/**
 * 목록 항목 전체 조회 (READ)
 */
async function getItems() {
    try {
        showStatus('데이터 불러오는 중...', 'info');
        var result = await graphCall(
            CONFIG.graphUrl + '/sites/' + APP.siteId + '/lists/' + APP.listId +
            '/items?$expand=fields,createdByUser&$top=500'
        );
        APP.data = result.value;
        renderTable();
        showStatus('로드 완료 (' + APP.data.length + '건)', 'success');
        document.getElementById('countInfo').innerText = '총 ' + APP.data.length + '건';
    } catch (e) {
        showStatus('로드 실패: ' + e.message, 'error');
    }
}

/**
 * 새 항목 생성 (CREATE)
 */
async function createItem(name, dept, pos) {
    var body = { fields: { Title: name } };
    body.fields['_xbd80__xc11c_'] = dept;
    body.fields['_xc9c1__xae09_'] = pos;
    await graphCall(
        CONFIG.graphUrl + '/sites/' + APP.siteId + '/lists/' + APP.listId + '/items',
        'POST', body
    );
}

/**
 * 항목 수정 (UPDATE)
 */
async function updateItem(id, name, dept, pos) {
    var body = { Title: name };
    body['_xbd80__xc11c_'] = dept;
    body['_xc9c1__xae09_'] = pos;
    await graphCall(
        CONFIG.graphUrl + '/sites/' + APP.siteId + '/lists/' + APP.listId + '/items/' + id + '/fields',
        'PATCH', body
    );
}

/**
 * 항목 삭제 (DELETE)
 */
async function deleteItemFromList(id) {
    await graphCall(
        CONFIG.graphUrl + '/sites/' + APP.siteId + '/lists/' + APP.listId + '/items/' + id,
        'DELETE'
    );
}

/**
 * 저장 (추가 또는 수정)
 */
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

/**
 * 삭제 (UI에서 호출)
 */
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
