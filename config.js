const msalConfig = {
  auth: {
	  //アプリの登録でのアプリケーションIDとリダイレクトURIの設定（office365ymsl_githubwebsite）
    clientId: '2aa23749-177c-4650-be79-a7f92385eac1',
    redirectUri: 'https://office365ymslchatbot.azurewebsites.net'
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
    forceRefresh: false
  }
};

const loginRequest = {
  scopes: [
    'openid',
    'profile',
    'user.read',
    'calendars.read'
  ]
}
