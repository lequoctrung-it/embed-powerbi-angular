import { Component, OnInit} from '@angular/core';
import { UserAgentApplication, AuthError, AuthResponse } from 'msal'
import { models } from 'powerbi-client'

import * as config from '../Config'
import {EventHandler} from "powerbi-client-angular/components/powerbi-embed/powerbi-embed.component";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  embedConfiguration: any

  ngOnInit() {
    this.authenticate();
  }

  authenticate(): void {
    const thisObj = this;

    const msalConfig = {
      auth: {
        clientId: config.clientId
      }
    };

    const loginRequest = {
      scopes: config.scopes
    }

    const msalInstance: UserAgentApplication = new UserAgentApplication(msalConfig);

    function successCallBack(response: AuthResponse): void {
      if (response.tokenType === "id_token") {
        thisObj.authenticate();
      } else if (response.tokenType === "access_token") {

        // Refresh User Permissions
        thisObj.tryRefreshUserPermissions(response.accessToken);
        thisObj.getembedUrl(response.accessToken);
      } else {
        console.error("Token type is: " + response.tokenType);
      }
    }

    function failCallBack(error: AuthError): void {
      console.error("Redirect error: " + error)
    }

    msalInstance.handleRedirectCallback(successCallBack, failCallBack);

    // Check if there is a cached user
    if (msalInstance.getAccount()) {
      // get access token silently from cached id-token
      msalInstance.acquireTokenSilent(loginRequest)
        .then((response: AuthResponse) => {
          // get access token from response: response.accessToken
          this.getembedUrl(response.accessToken)
        })
        .catch((err: AuthError) => {
          // refresh access token silently from cached id-token
          // makes the call to handleredirectcallback
          if (err.name === "InteractionRequiredAuthError") {
            msalInstance.acquireTokenRedirect(loginRequest);
          } else {
            console.error("Error: " + err.toString())
          }
        });
    } else {
      // user is not logged in or cached, you will need to log them in to acquire a token
      msalInstance.loginRedirect(loginRequest)
    }
  }

  // Power BI REST API call to refresh User Permissions in Power BI
  // Refreshes user permissions and makes sure the user permissions are fully updated
  // https://docs.microsoft.com/rest/api/power-bi/users/refreshuserpermissions
  tryRefreshUserPermissions(accessToken: string) {
    fetch("https://api.powerbi.com/v1.0/myorg/RefreshUserPermissions", {
      headers: {
        "Authorization": "Bearer " + accessToken
      },
      method: "POST"
    })
      .then(response => {
        if (response.ok) {
          console.log("User permissions refreshed successfully");
        } else {
          // Too many requests in one hour will cause the API to fail
          if (response.status === 429) {
            console.error("Permissions refresh will be available in up to an hour.");
          } else {
            console.error(response);
          }
        }
      })
      .catch(error => {
        console.error("Failure in making API call." + error)
      })
  }

  // Power BI REST API call to get the embed URL of the report
  getembedUrl(accessToken: string) {
    const thisObj = this;

    fetch("https://api.powerbi.com/v1.0/myorg/groups/" + config.workspaceId + "/reports/" + config.reportId, {
      headers: {
        "Authorization": "Bearer " + accessToken
      },
      method: "GET"
    })
      .then(response => {
        response.json()
          .then(body => {
            // Successful response
            if (response.ok) {
              this.embedPowerBI(accessToken, body["embedUrl"])
            } else {
              console.log("Error occured while fetching the embed URL of the report");
              console.log("Request Id: " + response.headers.get("requestId"));
              console.log("Error " + response.status + ": " + body.error.code);
            }
          })
          .catch(() => {
            console.log("Error occured while fetching the embed URL of the report");
            console.log("Request Id: " + response.headers.get("requestId"));
            console.log("Error " + response.status + ":  An error has occurred");
          })
      })
      .catch(error => {
        console.error(error)
      })
  }

  embedPowerBI(accessToken: string, embedUrl: string): void {
    // console.log("access token: " + accessToken, "embed Url: " + embedUrl)
    this.embedConfiguration = {
      type: "report",
      id: config.reportId,
      embedUrl: embedUrl,
      accessToken: accessToken,
      tokenType: models.TokenType.Aad,
      settings: {
        panes: {
          filters: {
            expanded: false,
            visible: false
          }
        },
        background: models.BackgroundType.Transparent,
      }
    }
  }

  eventHandler(): Map<string, EventHandler | null> {
    return new Map([
      ['loaded', () => console.log('Report loaded')],
      ['rendered', () => console.log('Report rendered')],
      ['error', (event: any) => console.log(event.detail)]
  ])
  }
}
