<?php
/**
 * Geni OAuth callback for the Holocaust projects export.
 *
 * Deploy: create a folder named "callback" in the schoenberg.com webroot
 * and upload this file into it AS "index.php", so that it answers at
 *     https://schoenberg.com/callback
 * (which is the Callback URL registered with the Geni app).
 *
 * What it does:
 *  - When Geni redirects here after you approve (?code=...), it stores the
 *    code in a file next to this script and shows a "you're done" page.
 *  - The export script polls this same URL with ?fetch=<key> and receives
 *    the stored code once (the code is deleted after pickup, and ignored
 *    if older than 9 minutes, since Geni codes expire in ~10).
 *
 *  - It also relays notification emails for the export script (POST with
 *    notify=1), using the server's mail() the same way the Virtual Arnold
 *    daily digest does — no SMTP credentials needed anywhere.
 *
 * $FETCH_KEY must match "auth_fetch_key" in geni_config.json.
 */

$FETCH_KEY = "CHANGE_ME_generate_a_long_random_key";  // must match auth_fetch_key in geni_config.json
$NOTIFY_TO = "you@example.com";
$NOTIFY_FROM = "Geni Export <geni-export@your-domain.example>";
$EXPORT_DIR = "/home/YOUR_USER/geni_export/exports";
$STORE = __DIR__ . "/geni_code_store.txt";
$MAX_AGE_SECONDS = 540;

if (isset($_GET["code"])) {
    // Redirect from Geni: store the code for the export script to pick up.
    $code = preg_replace("/[^A-Za-z0-9_\\-]/", "", $_GET["code"]);
    if ($code !== "") {
        file_put_contents($STORE, $code . "|" . time(), LOCK_EX);
    }
    header("Content-Type: text/html; charset=utf-8");
    echo "<!doctype html><html><head><title>Geni export authorized</title>";
    echo "<style>body{font-family:Georgia,serif;max-width:34em;margin:12vh auto;";
    echo "padding:0 1em;color:#333}h2{color:#1a5276}</style></head><body>";
    echo "<h2>Authorization received &#10003;</h2>";
    echo "<p>The Geni Holocaust projects export will pick this up and resume ";
    echo "automatically within a minute or two.</p>";
    echo "<p>You can close this window.</p></body></html>";
} elseif (isset($_POST["notify"], $_POST["fetch"])
          && is_string($_POST["fetch"])
          && hash_equals($FETCH_KEY, $_POST["fetch"])) {
    // Email relay for the export script (same mechanism as the Virtual
    // Arnold daily digest: the web server's own mail()).
    header("Content-Type: text/plain; charset=utf-8");
    $subject = isset($_POST["subject"]) ? trim($_POST["subject"]) : "Geni export";
    $subject = preg_replace("/[\\r\\n]+/", " ", mb_substr($subject, 0, 200));
    $body = isset($_POST["body"]) ? (string) $_POST["body"] : "";
    $headers = "From: " . $NOTIFY_FROM . "\r\n"
             . "Content-Type: text/plain; charset=UTF-8\r\n";
    echo @mail($NOTIFY_TO, $subject, $body, $headers) ? "OK" : "MAIL FAILED";
} elseif (isset($_GET["download"], $_GET["fetch"])
          && is_string($_GET["fetch"])
          && hash_equals($FETCH_KEY, $_GET["fetch"])) {
    // Secure download link for finished exports (used in the completion
    // email). Streams the file from the private export folder; only
    // export-named files can be requested.
    $name = basename((string) $_GET["download"]);
    if (!preg_match('/^[A-Za-z0-9_-]+\\.(csv|xlsx)$/', $name)) {
        http_response_code(400); exit("Bad file name.");
    }
    $path = $EXPORT_DIR . "/" . $name;
    if (!is_file($path)) {
        http_response_code(404);
        exit("Not found. The export may have been moved or renamed.");
    }
    $type = substr($name, -5) === ".xlsx"
        ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        : "text/csv";
    header("Content-Type: " . $type);
    header("Content-Length: " . filesize($path));
    header('Content-Disposition: attachment; filename="' . $name . '"');
    readfile($path);
} elseif (isset($_GET["fetch"]) && is_string($_GET["fetch"])
          && hash_equals($FETCH_KEY, $_GET["fetch"])) {
    // Poll from the export script: hand over the code once, then delete it.
    header("Content-Type: text/plain; charset=utf-8");
    if (file_exists($STORE)) {
        $parts = explode("|", trim(file_get_contents($STORE)), 2);
        unlink($STORE);
        if (count($parts) === 2 && (time() - (int)$parts[1]) < $MAX_AGE_SECONDS) {
            echo $parts[0];
        }
    }
    // Empty body = nothing waiting.
} else {
    http_response_code(204);
}
