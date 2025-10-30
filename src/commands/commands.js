/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // Office is ready.
});

// Minimal command action for the Ribbon button. Completes immediately.
function action(event) {
  // No-op for now — ribbon button opens the taskpane via manifest's ShowTaskpane action.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
