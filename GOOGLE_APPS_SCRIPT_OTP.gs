function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents || '{}');
    var to = payload.to;
    var name = payload.name || 'User';
    var otp = payload.otp || '';
    var subject = payload.subject || 'PitchCraft AI OTP Verification';
    var message = payload.message || ('Hello ' + name + '\n\nYour OTP is: ' + otp + '\n\nThis OTP will expire soon.');

    if (!to || !otp) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'Missing email or otp' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    GmailApp.sendEmail(to, subject, message, {
      name: 'PitchCraft AI',
      replyTo: Session.getActiveUser().getEmail()
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: 'OTP email sent' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
