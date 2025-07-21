  // Disable right-click context menu
    document.addEventListener("contextmenu", function (e) {
        e.preventDefault();
    }, false);

    // Disable common shortcuts for viewing source
    document.addEventListener("keydown", function (e) {
        // F12
        if (e.key === "F12") {
            e.preventDefault();
        }
        // Ctrl+Shift+I or J
        if (e.ctrlKey && e.shiftKey && (e.key === "I" || e.key === "J" || e.key === "C")) {
            e.preventDefault();
        }
        // Ctrl+U or Ctrl+S
        if (e.ctrlKey && (e.key === "U" || e.key === "S")) {
            e.preventDefault();
        }
    }, false);