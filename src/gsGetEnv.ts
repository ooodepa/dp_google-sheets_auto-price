function gsGetEnv(): Record<string, string> {
  let env = {};
  const envFile = HtmlService.createHtmlOutputFromFile('env.html').getContent();

  if (!envFile) {
    return {};
  }

  const lines = envFile.split('\n');

  for (let i = 0; i < lines.length; ++i) {
    const currentLine = lines[i];
    if (currentLine.length == 0) {
      continue;
    }

    if (currentLine[0] === '#') {
      continue;
    }

    let isFoundEqual = false;
    for (let j = 0; j < currentLine.length; ++j) {
      if (currentLine[j] === '=') {
        isFoundEqual = true;
        break;
      }
    }

    if (!isFoundEqual) {
      continue;
    }

    const keyValue = currentLine.split('=');
    const key = keyValue[0];
    const value = `${keyValue.slice(1).join('=')}`;

    env[key] = value;
  }

  return env;
}
