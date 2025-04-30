/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { getUserData } from "../helpers/sso-helper";
import { forwardAndDelete } from "../helpers/sso-helper";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("getProfileButton").onclick = run;
    document.getElementById("forwardAndDeleteButton").onclick = runForwardAndDelete;
  }
});

export function showConfirmationDialog(): void {
  // Display the confirmation dialog
  document.getElementById("confirmationDialog").classList.remove("hidden");

  // Set up event listeners for the dialog buttons
  document.getElementById("confirmButton").onclick = () => {
    document.getElementById("confirmationDialog").classList.add("hidden");
    runForwardAndDelete();
  };

  document.getElementById("cancelButton").onclick = () => {
    document.getElementById("confirmationDialog").classList.add("hidden");
  };
}

export async function runForwardAndDelete() {
  const messageId = Office.context.mailbox.item.itemId;
  const forwardToAddress = "carleid@carleid.onmicrosoft.com";
  forwardAndDelete(messageId, forwardToAddress, notifyUser);
}

export function notifyUser(): void {
  Office.context.mailbox.item.notificationMessages.addAsync("forwardAndDelete", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Forward and Delete operation completed",
    icon: "iconid",
    persistent: true,
  });
}

export async function run() {
  getUserData(writeDataToOfficeDocument);
}

export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];
  let userProfileInfo: string[] = [];
  userProfileInfo.push(result["displayName"]);
  userProfileInfo.push(result["jobTitle"]);
  userProfileInfo.push(result["mail"]);
  userProfileInfo.push(result["mobilePhone"]);
  userProfileInfo.push(result["officeLocation"]);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }
  Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
}
