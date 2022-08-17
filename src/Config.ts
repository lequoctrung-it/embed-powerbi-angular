// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

/* eslint-disable @typescript-eslint/no-inferrable-types */

// Scope of AAD app. Use the below configuration to use all the permissions provided in the AAD app through Azure portal.
// Refer https://aka.ms/PowerBIPermissions for complete list of Power BI scopes
export const scopes: string[] = ["https://analysis.windows.net/powerbi/api/Report.Read.All"];

// Client Id (Application Id) of the AAD app.
export const clientId: string = "dfcd308e-6439-4823-b1c0-8753d176bd52";

// Id of the workspace where the report is hosted
export const workspaceId: string = "91f97e7c-c0ee-4eb8-a997-07fc9226c461";

// Id of the report to be embedded
export const reportId: string = "6356ed9b-6e98-4d41-9338-52b4601da274";
