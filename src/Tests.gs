/**
 * Simulate the interview alert text and routing without creating an event or contacting Twilio.
 * @requires Initialized tracker with MODE=DRY_RUN.
 * @effects Writes the chosen channels and generic message to the Apps Script execution log only.
 * @throws {Error} If the tracker is not in DRY_RUN or settings cannot be read. 
 */
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
