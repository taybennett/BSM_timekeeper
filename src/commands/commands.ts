/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

const API = "https://mahfk-api-production.up.railway.app";

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function validateBeforeSend(event: any) {
  let isCompleted = false;

  const completeSend = (allowEvent: boolean, errorMessage?: string) => {
    if (isCompleted) {
      return;
    }

    isCompleted = true;
    event.completed(errorMessage ? { allowEvent, errorMessage } : { allowEvent });
  };

  const getSavedUser = () => {
    try {
      const savedUser = localStorage.getItem("mahfk_user");
      return savedUser ? JSON.parse(savedUser) : null;
    } catch (error) {
      return null;
    }
  };

  const getSubject = (item: any, callback: (subject: string) => void) => {
    if (typeof item.subject === "string") {
      callback(item.subject);
      return;
    }

    if (item.subject && typeof item.subject.getAsync === "function") {
      item.subject.getAsync((result: any) => {
        callback(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "");
      });
      return;
    }

    callback("");
  };

  const getBody = (item: any, callback: (body: string) => void) => {
    if (!item.body || typeof item.body.getAsync !== "function") {
      callback("");
      return;
    }

    item.body.getAsync(Office.CoercionType.Text, (result: any) => {
      callback(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "");
    });
  };

  const saveHoursToItem = (
    hours: number,
    submittedAt: string,
    matterId: string,
    matterName: string,
    callback: (success: boolean) => void
  ) => {
    const item = Office.context.mailbox.item;

    item.loadCustomPropertiesAsync((loadResult: any) => {
      if (loadResult.status !== Office.AsyncResultStatus.Succeeded) {
        callback(false);
        return;
      }

      const customProps = loadResult.value;
      customProps.set("mahfk_hours", hours.toString());
      customProps.set("mahfk_hours_submitted_at", submittedAt);
      customProps.set("mahfk_matter_id", matterId || "");
      customProps.set("mahfk_matter_name", matterName || "");

      customProps.saveAsync((saveResult: any) => {
        callback(saveResult.status === Office.AsyncResultStatus.Succeeded);
      });
    });
  };

  const syncHoursToApi = (hours: number, submittedAt: string, matterId: string, matterName: string) => {
    const item = Office.context.mailbox.item;
    const user = getSavedUser();

    getSubject(item, (subject) => {
      getBody(item, (body) => {
        fetch(`${API}/entries/from-email`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            subject: subject,
            body: body.slice(0, 1500),
            user_id: user && user.id ? user.id : null,
            matter_id: matterId || undefined,
            hours: hours,
            source: "outlook-send-dialog",
            submitted_at: submittedAt,
          }),
        })
          .then((response) => {
            if (!response.ok) {
              throw new Error("MAHFK entry creation failed");
            }

            return response.json();
          })
          .then((entry) => {
            if (!entry || !entry.id) {
              return entry;
            }

            return fetch(`${API}/entries/${entry.id}`, {
              method: "PUT",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                ...entry,
                matter_id: matterId || entry.matter_id,
                hours: hours,
              }),
            }).then((response) => {
              if (!response.ok) {
                throw new Error("MAHFK entry update failed");
              }

              return response.json().catch(() => entry);
            });
          })
          .then(() => {
            saveHoursToItem(hours, submittedAt, matterId, matterName, (success) => {
              if (success) {
                completeSend(true);
                return;
              }

              completeSend(false, "MAHFK entry saved, but Outlook hours could not be stored. Please try again.");
            });
          })
          .catch(() => {
            completeSend(false, "Could not save hours to MAHFK. Please try again.");
          });
      });
    });
  };

  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/assets/Mahfk.html",
    { height: 62, width: 38, displayInIframe: true },
    function (asyncResult: any) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg: any) {
          try {
            const message = JSON.parse(arg.message);

            if (message.action === "submit" && Number(message.hours) > 0) {
              dialog.close();
              syncHoursToApi(
                Number(message.hours),
                message.submittedAt || new Date().toISOString(),
                message.matterId || "",
                message.matterName || ""
              );
              return;
            }

            dialog.close();
            completeSend(false, "Send cancelled. Please submit billable hours before sending.");
          } catch (error) {
            dialog.close();
            completeSend(false, "Invalid MAHFK hours selection. Please try again.");
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
          completeSend(false, "Send cancelled. Please submit billable hours before sending.");
        });
      } else {
        completeSend(false, "MAHFK hours dialog failed to open. Please try again.");
      }
    }
  );
}

// Register the function with Office.
Office.actions.associate("validateBeforeSend", validateBeforeSend);
