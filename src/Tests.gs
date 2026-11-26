/** Preview the interview message and route without contacting Twilio. */
function simulateInterviewAlert() {
  const config = getSettings_();
  if (config.MODE !== 'DRY_RUN') throw Error('Switch to DRY_RUN first.');
  Logger.log([
    'SIMULATED INTERVIEW INVITATION',
    'Mode: ' + config.MODE,
    'Channels: ' + channels_(config.ROUTE_INTERVIEW).join(', '),
    '',
    'Message: ' + notificationText_('INTERVIEW'),
    '',
    'No SMS or phone calls were initiated.'
  ].join('\n'));
}
