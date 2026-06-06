// =============================================
// GITHUB.gs - Push deployed JSON to GitHub
// =============================================

// --------------------------------------------
// DEPLOY: Build payload and push to GitHub
// Called from the CMS Manager menu
// --------------------------------------------
function deployToGithub() {
  var ui = SpreadsheetApp.getUi();

  try {
    var token    = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    if (!token)  throw new Error('GITHUB_TOKEN is not set in Script Properties.');

    var payload  = buildDeployedPayload();
    var json     = JSON.stringify(payload, null, 2);
    var encoded  = Utilities.base64Encode(json);

    var apiUrl   = 'https://api.github.com/repos/'
                 + CONFIG.GITHUB_OWNER + '/'
                 + CONFIG.GITHUB_REPO
                 + '/contents/'
                 + CONFIG.GITHUB_FILE_PATH;

    // Check if the file already exists so we can pass the required SHA
    var sha      = getGithubFileSha(apiUrl, token);

    var body     = {
      message: 'CMS deploy: ' + new Date().toISOString(),
      branch:  CONFIG.GITHUB_BRANCH,
      content: encoded
    };

    if (sha) body.sha = sha; // Required by GitHub API to update an existing file

    var options  = {
      method:      'PUT',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Accept':        'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28'
      },
      payload:     JSON.stringify(body),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(apiUrl, options);
    var code     = response.getResponseCode();
    var result   = JSON.parse(response.getContentText());

    if (code === 200 || code === 201) {
      var fileUrl = result.content && result.content.html_url
                  ? result.content.html_url
                  : apiUrl;

      ui.alert(
        'Deploy Successful',
        'JSON pushed to GitHub.\n\nFile: ' + CONFIG.GITHUB_FILE_PATH
        + '\nBranch: ' + CONFIG.GITHUB_BRANCH
        + '\nCommit: ' + (result.commit && result.commit.sha
                          ? result.commit.sha.substring(0, 7)
                          : 'n/a'),
        ui.ButtonSet.OK
      );

    } else {
      throw new Error('GitHub API returned ' + code + ': '
                    + (result.message || response.getContentText()));
    }

  } catch (err) {
    ui.alert('Deploy Failed', err.message, ui.ButtonSet.OK);
    Logger.log('deployToGithub error: ' + err.message);
  }
}

// --------------------------------------------
// HELPER: Fetch the current SHA of the file
// GitHub requires this to update an existing file
// --------------------------------------------
function getGithubFileSha(apiUrl, token) {
  try {
    var options  = {
      method:  'GET',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Accept':        'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28'
      },
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(apiUrl, options);
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      return data.sha || null;
    }
    return null; // File does not exist yet

  } catch (err) {
    Logger.log('getGithubFileSha error: ' + err.message);
    return null;
  }
}

// --------------------------------------------
// PREVIEW: Log the payload without pushing
// Run from the Apps Script editor to test
// --------------------------------------------
function previewDeployPayload() {
  var payload = buildDeployedPayload();
  Logger.log(JSON.stringify(payload, null, 2));
  SpreadsheetApp.getUi().alert(
    'Payload Preview',
    'Check the Apps Script Logs for the full JSON output.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}