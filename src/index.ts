/// <reference types="office-js" />

export const showLoginPopup = (popupHTML: string) => {
  const url = window.location.origin + popupHTML;

  Office.context.ui.displayDialogAsync(
    url,
    { height: 70, width: 28 },
    (result) => {
      const dialog = result.value;

      // Add event handler for the 'messageParent' event
      dialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        async (args) => {
          console.log(url);
        }
      );
    }
  );
};
