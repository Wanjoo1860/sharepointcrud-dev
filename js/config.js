// ============================================================
// ★★★ 설정 값 ★★★
// ============================================================
var CONFIG = {
    clientId: '4aeb92a9-8ef6-476a-a419-9125032309fd',
    tenantId: 'cc13b6f1-ef21-479f-9853-3e2dffa71d6b',
    siteHostname: 'globalsoft.sharepoint.com',
    sitePath: '/sites/Dev896',
    listName: 'TestData',
    adminGroupId: 'bc9227b3-a99a-459e-bfc7-71c7166f19c4',
    scopes: ['User.Read', 'Sites.ReadWrite.All', 'GroupMember.ReadWrite.All'],
    graphUrl: 'https://graph.microsoft.com/v1.0'
};

// ============================================================
// ★★★ 전역 상태 ★★★
// ============================================================
var APP = {
    accessToken: null,
    msalInstance: null,
    isInTeams: false,
    msalReady: false,
    siteId: null,
    listId: null,
    data: [],
    currentUserRole: 'user',
    currentUserEmail: '',
    currentUserId: '',
    adminMembers: []
};
