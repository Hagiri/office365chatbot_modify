// Create an options object with the same scopes from the login
const options =
  new MicrosoftGraph.MSALAuthenticationProviderOptions([
    'user.read',
    'calendars.read'
  ]);
// Create an authentication provider for the implicit flow
const authProvider =
  new MicrosoftGraph.ImplicitMSALAuthenticationProvider(msalClient, options);
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({authProvider});


async function getEvents() {
  try {
    let events = await graphClient
        .api('/me/events')
        .select('subject,organizer,start,end,location')
        .orderby('createdDateTime DESC')
        .get();

    updatePage(msalClient.getAccount(), Views.calendar, events);
  } catch (error) {
    updatePage(msalClient.getAccount(), Views.error, {
      message: 'Error getting events',
      debug: error
    });
  }
}

/* async function getEvents() {
  //add for showimg profile
  try {
    let events = await graphClient
        .api('/me')
        .select('displayName,surname,givenName,mail,userPrincipalName,id')
        .get();

    updatePage(msalClient.getAccount(), Views.calendar, Views.profile, events);
  } catch (error) {
    updatePage(msalClient.getAccount(), Views.error, {
      message: 'Error getting profile',
      debug: error
    });
  }
}   */
  
