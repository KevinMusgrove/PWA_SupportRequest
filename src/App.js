import React, { Component } from 'react';
import { BrowserRouter as Router, Route } from 'react-router-dom';
import { Container } from 'reactstrap';
import NavBar from './NavBar';
import ErrorMessage from './ErrorMessage';
import Welcome from './Welcome';
import 'bootstrap/dist/css/bootstrap.css';
import config from './Config';
import { UserAgentApplication } from 'msal';
import { getUserDetails, getUserBlob, getUserManager } from './GraphService';

class App extends Component {
  constructor(props) {
    super(props);
  
    this.userAgentApplication = new UserAgentApplication({
          auth: {
              clientId: config.appId
          },
          cache: {
              cacheLocation: "localStorage",
              storeAuthStateInCookie: true
          }
      });
    var user = this.userAgentApplication.getAccount();
  
    this.state = {
      isAuthenticated: (user !== null),
      user: {},
      files:{},
      avatar:'',
      managerEmail:'',
      error: null,
    };
  
    if (user) {
      // Enhance user object with data from Graph
      this.getUserProfile();   

    }
  }

  async login() {
    try {
      await this.userAgentApplication.loginPopup(
          {
            scopes: config.scopes,
            prompt: "select_account"
        });
      await this.getUserProfile();
    }
    catch(err) {
      //debugger;
      if(err.errorCode !== "user_cancelled")
      {
        var errParts = err.split('|');
        this.setState({
          isAuthenticated: false,
          user: {},
          avatar:'',
          files:{},
          error: { message: errParts[1], debug: errParts[0] }
        });
      }
    }
  }

  logout() {
    this.userAgentApplication.logout();
  }

  async getUserProfile() {
    try {
      // Get the access token silently
      // If the cache contains a non-expired token, this function
      // will just return the cached token. Otherwise, it will
      // make a request to the Azure OAuth endpoint to get a token
      var accessToken = await this.userAgentApplication.acquireTokenSilent({
        scopes: config.scopes
      });
  
      if (accessToken) {
        // Get the user's profile from Graph
        var user = await getUserDetails(accessToken);
        var managerEmail = await getUserManager(accessToken);
        var blobImage = await getUserBlob(accessToken.accessToken);        
        if(blobImage){
          var reader = new FileReader();
          reader.readAsDataURL(blobImage); 
          reader.onloadend = (e) => {
              var base64data = reader.result;
              this.setState({
                avatar:base64data
              });
           }
        }                   
        this.setState({
          isAuthenticated: true,
          user: {
            firstName: user.givenName,
            lastName: user.surname,
            displayName: user.displayName,
            email: user.mail || user.userPrincipalName,
            cellPhone: user.mobilePhone,
            bussinessPhone: user.businessPhones["0"] || null
          },
          error: null,
          managerEmail:managerEmail
        });
      }
    }
    catch(err) {
      var error = {};
      if (typeof(err) === 'string') {
        var errParts = err.split('|');
        error = errParts.length > 1 ?
          { message: errParts[1], debug: errParts[0] } :
          { message: err };
      } else {
        error = {
          message: err.message,
          debug: JSON.stringify(err)
        };
      }
  
      this.setState({
        isAuthenticated: false,
        user: {},
        error: error
      });
    }
  }          

  render() {
    let error = null;
    if (this.state.error) {
      error = <ErrorMessage message={this.state.error.message} debug={this.state.error.debug} />;
    }

    return (
      <Router>
        <div>
        <NavBar
          isAuthenticated={this.state.isAuthenticated}
          authButtonMethod={this.state.isAuthenticated ? this.logout.bind(this) : this.login.bind(this)}
          user={this.state.user} avatar={this.state.avatar}/>
          <Container>
            {error}
            <Route exact path="/"
              render={(props) =>
            <Welcome {...props}
              isAuthenticated={this.state.isAuthenticated}
              UserAgentApplication={this.userAgentApplication}
              user={this.state.user}
              managerEmail={this.state.managerEmail}
              files={this.state.files}
              authButtonMethod={this.login.bind(this)} />
              } />
          </Container>
        </div>
      </Router>
    );
  }

  setErrorMessage(message, debug) {
    this.setState({
      error: {message: message, debug: debug}
    });
  }
}

export default App;