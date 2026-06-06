// =============================================
// GITHUB.gs - Push deployed JSON to GitHub
// =============================================

function deployToGithub() {
  var ui = SpreadsheetApp.getUi();

  try {
    var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    if (!token) throw new Error('GITHUB_TOKEN is not set in Script Properties.');

    var payload = buildDeployedPayload();
    var json    = JSON.stringify(payload, null, 2);
    var encoded = Utilities.base64Encode(Utilities.newBlob(json).getBytes());

    var apiUrl  = 'https://api.github.com/repos/'
                + CONFIG.GITHUB_OWNER + '/'
                + CONFIG.GITHUB_REPO
                + '/contents/'
                + CONFIG.GITHUB_FILE_PATH;

    var sha     = getGithubFileSha(apiUrl, token);

    var body    = {
      message: 'CMS deploy: ' + new Date().toISOString(),
      branch:  CONFIG.GITHUB_BRANCH,
      content: encoded
    };
    if (sha) body.sha = sha;

    var options = {
      method:      'PUT',
      contentType: 'application/json',
      headers:     buildHeaders(token),
      payload:     JSON.stringify(body),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(apiUrl, options);
    var code     = response.getResponseCode();
    var result   = JSON.parse(response.getContentText());

    if (code === 200 || code === 201) {
      ui.alert(
        'Deploy Successful',
        'data/cms.json pushed to GitHub.\n\n'
        + 'Branch: ' + CONFIG.GITHUB_BRANCH + '\n'
        + 'Commit: ' + (result.commit && result.commit.sha
                         ? result.commit.sha.substring(0, 7) : 'n/a'),
        ui.ButtonSet.OK
      );
    } else if (code === 403) {
      throw new Error(
        'Permission denied (403).\n\n'
        + 'Your token needs "Contents: Read and write" permission.\n\n'
        + 'Fix options:\n'
        + '1. Classic token: make sure the "repo" scope is checked.\n'
        + '2. Fine-grained token: go to Settings → Developer settings → '
        + 'Fine-grained tokens → edit your token → Repository permissions '
        + '→ Contents → Read and write.\n\n'
        + 'GitHub message: ' + (result.message || '')
      );
    } else if (code === 404) {
      throw new Error(
        'Repo not found (404). Check GITHUB_OWNER and GITHUB_REPO in Config.js.\n'
        + 'Current values: ' + CONFIG.GITHUB_OWNER + '/' + CONFIG.GITHUB_REPO + '\n\n'
        + 'GitHub message: ' + (result.message || '')
      );
    } else {
      throw new Error('GitHub API returned ' + code + ': ' + (result.message || response.getContentText()));
    }

  } catch (err) {
    ui.alert('Deploy Failed', err.message, ui.ButtonSet.OK);
    Logger.log('deployToGithub error: ' + err.message);
  }
}

// --------------------------------------------
// HELPER: Fetch current SHA (needed to update an existing file)
// --------------------------------------------
function getGithubFileSha(apiUrl, token) {
  try {
    var response = UrlFetchApp.fetch(apiUrl, {
      method:  'GET',
      headers: buildHeaders(token),
      muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText()).sha || null;
    }
    return null;
  } catch (err) {
    Logger.log('getGithubFileSha error: ' + err.message);
    return null;
  }
}

// --------------------------------------------
// HELPER: Build auth headers — uses "token" prefix
// which works for both classic and fine-grained PATs
// --------------------------------------------
function buildHeaders(token) {
  return {
    'Authorization':        'token ' + token,
    'Accept':               'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28',
    'User-Agent':           'JPAR-CMS-Script'
  };
}

// --------------------------------------------
// PREVIEW: Log payload without pushing (debug)
// --------------------------------------------
function previewDeployPayload() {
  Logger.log(JSON.stringify(buildDeployedPayload(), null, 2));
  SpreadsheetApp.getUi().alert(
    'Payload Preview',
    'Check the Apps Script Logs for the full JSON output.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
