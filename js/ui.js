// ============================================================
// UI 렌더링 / 이벤트 처리
// ============================================================

/**
 * 데이터 테이블 렌더링
 */
function renderTable() {
    var tbody = document.getElementById('tableBody');
    if (APP.data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:30px;color:#999;">데이터가 없습니다.</td></tr>';
        return;
    }

    var html = '';
    APP.data.forEach(function (item) {
        var f = item.fields;
        var id = f.id || item.id;

        html += '<tr>';
        html += '<td>' + id + '</td>';
        html += '<td>' + (f.Title || '') + '</td>';
        html += '<td>' + (f['_xbd80__xc11c_'] || '') + '</td>';
        html += '<td>' + (f['_xc9c1__xae09_'] || '') + '</td>';
        html += '<td>';

        if (APP.currentUserRole === 'admin') {
            // 관리자: 모든 항목 수정/삭제 가능
            html += '<button class="btn btn-warning" onclick="editItem(\'' + id + '\')">수정</button> ';
            html += '<button class="btn btn-danger" onclick="deleteItem(\'' + id + '\')">삭제</button>';
        } else {
            // 일반 사용자: 본인 항목만 수정/삭제
            var createdByEmail = '';
            if (item.createdBy && item.createdBy.user && item.createdBy.user.email) {
                createdByEmail = item.createdBy.user.email.toLowerCase();
            }
            var isOwner = (APP.currentUserEmail.toLowerCase() === createdByEmail);

            if (isOwner) {
                html += '<button class="btn btn-warning" onclick="editItem(\'' + id + '\')">수정</button> ';
                html += '<button class="btn btn-danger" onclick="deleteItem(\'' + id + '\')">삭제</button>';
            } else {
                html += '<span style="color:#999;font-size:12px;">읽기 전용</span>';
            }
        }

        html += '</td>';
        html += '</tr>';
    });
    tbody.innerHTML = html;
}

/**
 * 수정 모드 진입
 */
function editItem(id) {
    var item = APP.data.find(function (i) { return (i.fields.id || i.id) == id; });
    if (!item) return;
    var f = item.fields;
    document.getElementById('editId').value = id;
    document.getElementById('inputName').value = f.Title || '';
    document.getElementById('inputDept').value = f['_xbd80__xc11c_'] || '';
    document.getElementById('inputPosition').value = f['_xc9c1__xae09_'] || '';
    document.getElementById('formTitle').innerText = '데이터 수정 (ID: ' + id + ')';
}

/**
 * 수정 취소 / 폼 초기화
 */
function cancelEdit() {
    document.getElementById('editId').value = '';
    document.getElementById('inputName').value = '';
    document.getElementById('inputDept').value = '';
    document.getElementById('inputPosition').value = '';
    document.getElementById('formTitle').innerText = '데이터 추가';
}

/**
 * 상태 메시지 표시
 * @param {string} msg - 표시할 메시지
 * @param {string} type - 'success', 'error', 'info'
 */
function showStatus(msg, type) {
    var el = document.getElementById('status');
    el.innerText = msg;
    el.className = 'status show ' + type;
    if (type !== 'info') {
        setTimeout(function () { el.className = 'status'; }, 4000);
    }
}
