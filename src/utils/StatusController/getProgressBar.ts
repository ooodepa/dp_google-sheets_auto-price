function getProgressBar(percentage: number, barLength: number = 60) {
  const completedLength = Math.floor((percentage / 100) * barLength);
  const progressBar =
    '[' +
    'x'.repeat(completedLength) +
    '-'.repeat(barLength - completedLength) +
    '] ';
  return progressBar;
}
