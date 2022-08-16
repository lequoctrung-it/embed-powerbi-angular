import {AfterContentInit, Component, OnInit} from '@angular/core';
import { UserAgentApplication, AuthError, AuthResponse } from 'msal'
import { service, factories, models, IEmbedConfiguration } from 'powerbi-client'

import * as config from '../Config'
import {jsDocComment} from "@angular/compiler";

const powerbi = new service.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);

let accessToken = "";
let embedUrl = "";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements AfterContentInit, OnInit {
  ngOnInit() {
    if (this.state.accessToken !== "" && this.state.embedUrl !== "") {

      const embedConfiguration: IEmbedConfiguration = {
        type: "report",
        tokenType: models.TokenType.Aad,
        accessToken,
        embedUrl,
        id: config.reportId,
        /*
        // Enable this setting to remove gray shoulders from embedded report
        settings: {
            background: models.BackgroundType.Transparent
        }
        */
      };

      const report = powerbi.embed(reportContainer, embedConfiguration);

      // Clear any other loaded handler events
      report.off("loaded");

      // Triggers when a content schema is successfully loaded
      report.on("loaded", function () {
        console.log("Report load successful");
      });

      // Clear any other rendered handler events
      report.off("rendered");

      // Triggers when a content is successfully embedded in UI
      report.on("rendered", function () {
        console.log("Report render successful");
      });

      // Clear any other error handler event
      report.off("error");

      // Below patch of code is for handling errors that occur during embedding
      report.on("error", function (event) {
        const errorMsg = event.detail;

        // Use errorMsg variable to log error in any destination of choice
        console.error(errorMsg);
      });
  }

  ngAfterContentInit(): void {
    // User input - null check
    if (config.workspaceId === "" || config.reportId === "") {
      console.error("Please assign values to workspace Id and report Id in Config.ts file")
    } else {

      // Authenticate the user and generate the access token
      this.authenticate();
    }
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
        accessToken = response.accessToken;

        // Refresh User Permissions
        thisObj.tryRefreshUserPermissions();
        thisObj.getembedUrl();
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
          accessToken = response.accessToken
          this.getembedUrl()
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
  tryRefreshUserPermissions() {
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
  getembedUrl() {
    const thisObj = this;

    fetch("https://api.powerbi.com/v1.0/myorg/groups/" + config.workspaceId + "/reports/" + config.reportId, {
      headers: {
        "Authorization": "Bearer " + accessToken
      },
      method: "GET"
    })
      .then(response => {
        const errorMessage: string[] = [];
        errorMessage.push("Error occured while fetching the embed URL of the report");
        errorMessage.push("Request Id: " + response.headers.get("requestId"));

        response.json()
          .then(body => {
            // Successful response
            if (response.ok) {
              embedUrl = body["embedUrl"];
            } else {
              errorMessage.push("Error " + response.status + ": " + body.error.code);

              console.log(errorMessage);
            }
          })
          .catch(() => {
            errorMessage.push("Error " + response.status + ":  An error has occurred");

            console.log(errorMessage);
          })
      })
      .catch(error => {
        console.error(error)
      })
  }
}
