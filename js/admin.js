// ============================================================
// 관리자 역할 확인 / 관리자 추가 / 삭제
// ============================================================

/**
 * 현재 로그인한 사용자가 관리자 그룹의 구성원인지 확인
 */
async function checkUserRole() {
    try {
        // 1. 현재 사용자 ID 가져오기
        var meData = await graphGet(CONFIG.graphUrl + '/me');
        APP.currentUserId = meData.id;

        // 2. 관리자 그룹 멤버인지 확인
        var checkData = await graphGet(
            CONFIG.graphUrl + '/me/memberOf?$filter=id eq \'' + CONFIG.adminGroupId + '\''
        );

        if (checkData.value && checkData.value.length > 0) {
            APP.currentUserRole = 'admin';
            document.getElementById('loginStatus').innerHTML =
                '✓ ' + APP.currentUserEmail +
                ' <span class="role-badge role-admin">관리자</span>';
            document.getElementById('adminPanel').style.display = 'block';
            console.log('[권한] 관리자 확인됨');
            loadAdminMembers();
        } else {
            APP.currentUserRole = 'user';
            document.getElementById('loginStatus').innerHTML =
                '✓ ' + APP.currentUserEmail +
                ' <span class="role-badge role-user">사용자</span>';
            document.getElementById('adminPanel').style.display = 'none';
            console.log('[권한] 일반 사용자');
        }
    } catch (e) {
        console.warn('[권한] 역할 확인 실패:', e.message);
        APP.currentUserRole = 'user';
        document.getElementById('adminPanel').style.display = 'none';
    }
}

/**
 * 관리자 그룹의 전체 구성원 목록 조회
 */
async function loadAdminMembers() {
    try {
        var data = await graphGet(
            CONFIG.graphUrl + '/groups/' + CONFIG.adminGroupId + '/members?$select=id,displayName,mail,userPrincipalName'
        );
        APP.adminMembers = data.value || [];
        renderAdminTable();
    } catch (e) {
        showStatus('관리자 목록 조회 실패: ' + e.message, 'error');
    }
}

/**
 * 이메일로 사용자를 관리자 그룹에 추가
 */
async function addAdmin() {
    var email = document.getElementById('inputAdminEmail').value.trim();
    if (!email) {
        showStatus('추가할 사용자의 이메일을 입력하세요.', 'error');
        return;
    }

    try {
        // 1. 이메일로 사용자 조회
        var userResponse = await graphGet(CONFIG.graphUrl + '/users/' + email);
        var userId = userResponse.id;

        // 2. 그룹에 구성원으로 추가
        var addResponse = await graphPost(
            CONFIG.graphUrl + '/groups/' + CONFIG.adminGroupId + '/members/$ref',
            { '@odata.id': CONFIG.graphUrl + '/directoryObjects/' + userId }
        );

        if (addResponse.status === 204 || addResponse.ok) {
            showStatus(email + ' 관리자로 추가 완료', 'success');
            document.getElementById('inputAdminEmail').value = '';
            loadAdminMembers();
        } else if (addResponse.status === 400) {
            showStatus('이미 관리자로 등록된 사용자입니다.', 'error');
        } else {
            var errText = await addResponse.text();
            showStatus('추가 실패: ' + errText, 'error');
        }
    } catch (e) {
        if (e.message.indexOf('404') > -1) {
            showStatus('사용자를 찾을 수 없습니다: ' + email, 'error');
        } else {
            showStatus('관리자 추가 실패: ' + e.message, 'error');
        }
    }
}

/**
 * 관리자 그룹에서 구성원 제거
 * @param {string} userId - 제거할 사용자 ID
 * @param {string} displayName - 사용자 표시 이름 (확인용)
 */
async function removeAdmin(userId, displayName) {
    if (!confirm(displayName + ' 님을 관리자에서 제거하시겠습니까?')) return;

    // 자기 자신 제거 방지
    if (userId === APP.currentUserId) {
        showStatus('자기 자신은 제거할 수 없습니다.', 'error');
        return;
    }

    try {
        var response = await graphDelete(
            CONFIG.graphUrl + '/groups/' + CONFIG.adminGroupId + '/members/' + userId + '/$ref'
        );

        if (response.status === 204 || response.ok) {
            showStatus(displayName + ' 관리자에서 제거 완료', 'success');
            loadAdminMembers();
        } else {
            var errText = await response.text();
            showStatus('제거 실패: ' + errText, 'error');
        }
    } catch (e) {
        showStatus('관리자 제거 실패: ' + e.message, 'error');
    }
}

/**
 * 관리자 목록 테이블 렌더링
 */
function renderAdminTable() {
    var tbody = document.getElementById('adminTableBody');
    if (APP.adminMembers.length === 0) {
        tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;padding:20px;color:#999;">관리자가 없습니다.</td></tr>';
        return;
    }

    var html = '';
    APP.adminMembers.forEach(function (member) {
        var email = member.mail || member.userPrincipalName || '';
        var isSelf = (member.id === APP.currentUserId);
        html += '<tr>';
        html += '<td>' + (member.displayName || '') + '</td>';
        html += '<td>' + email + '</td>';
        html += '<td>';
        if (isSelf) {
            html += '<span style="color:#999;font-size:12px;">본인</span>';
        } else {
            html += '<button class="btn btn-danger" onclick="removeAdmin(\'' + member.id + '\', \'' + (member.displayName || '').replace(/'/g, "\\'") + '\')">제거</button>';
        }
        html += '</td>';
        html += '</tr>';
    });
    tbody.innerHTML = html;
}
