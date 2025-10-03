Office.onReady(() => {
  if (Office.context.host === Office.HostType.Excel) {
    openLoginDialog(); // Automatically open when Excel add-in loads
    console.log("Hello");
  }
});


function openLoginDialog() {
  const loginUrl = `https://auth.atlassian.com/authorize?audience=api.atlassian.com&client_id=EvglXZt0zxqDO4pNdRPxrPBiHESGWDqn&scope=read%3Ajira-user%20read%3Ajira-work%20offline_access&redirect_uri=http://localhost:3000/callback&response_type=code&prompt=consent`;

  Office.context.ui.displayDialogAsync(
    loginUrl,
    { height: 60, width: 40 },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Dialog failed to open:", asyncResult.error);
        return;
      }

      const dialog = asyncResult.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        try {
          const msg = JSON.parse(arg.message);
          dialog.close();

          if (msg.status === "ok") {
            console.log("✅ User logged in successfully");
            localStorage.setItem("authCode", msg.code);
          } else {
            console.error("❌ Login failed:", msg);
          }
        } catch (err) {
          console.error("Message parsing error:", err);
        }
      });
    }
  );
}
