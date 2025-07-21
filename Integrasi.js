const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const MIDTRANS_SERVER_KEY = SCRIPT_PROPS.getProperty("MIDTRANS_SERVER_KEY");
const SHEET_NAME = SCRIPT_PROPS.getProperty("SHEET_NAME");
const MIDTRANS_SNAP_API_URL = SCRIPT_PROPS.getProperty("MIDTRANS_SNAP_API_URL");

function doPost(e) {
  try {
    const requestBody = e.postData.contents;
    const payload = JSON.parse(requestBody);

    if (payload.action === "createMidtransPayment") { 
      createMidtransPaymentLink(payload);
      return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Payment link created"})).setMimeType(ContentService.MimeType.JSON);
    } else if (payload.transaction_status && payload.order_id) {
      handleMidtransNotification(payload);
      return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Notification received"})).setMimeType(ContentService.MimeType.JSON);
    } else {
      throw new Error("Unknown Request Type");
    }
  } catch (error) {
    throw new Error("Server Error: " + error.message);
  }
}

function handleMidtransNotification(appSheetPayload) {
    const fullOrderId = appSheetPayload.order_id;
    const orderId = fullOrderId.split("-")[0];
    const transactionStatus = appSheetPayload.transaction_status;
    const fraudStatus = appSheetPayload.fraud_status;
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const rows = sheet.getDataRange().getValues();
    const header = rows[0];
    const idxReservationId = header.indexOf("BookingID");
    const idxPaymentStatus = header.indexOf("PaymentStatus");
    const idxStatus = header.indexOf("Status");
    const idxBookedAt = header.indexOf("BookedAt");

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idxReservationId] === orderId) {
        if ((transactionStatus === "settlement" || transactionStatus === "capture") && (!appSheetPayload.hasOwnProperty("fraud_status") || fraudStatus === "accept")) {
          const now = new Date();
          sheet.getRange(i + 1, idxStatus + 1).setValue("Booked");
          sheet.getRange(i + 1, idxPaymentStatus + 1).setValue("Paid");
          sheet.getRange(i + 1, idxBookedAt + 1).setValue(now);
        } else if (transactionStatus === "pending") {
          sheet.getRange(i + 1, idxStatus + 1).setValue("Approved by Owner");
          sheet.getRange(i + 1, idxPaymentStatus + 1).setValue("Pending Guest Payment");
        } else if (transactionStatus === "cancel") {
          sheet.getRange(i + 1, idxStatus + 1).setValue("Cancelled by Guest");
          sheet.getRange(i + 1, idxPaymentStatus + 1).setValue("Cancelled by Guest");
        } else if (transactionStatus === "expire") {
          sheet.getRange(i + 1, idxStatus + 1).setValue("Cancelled by System");
          sheet.getRange(i + 1, idxPaymentStatus + 1).setValue("Payment Expired");
        } else if (transactionStatus === "deny") {
          sheet.getRange(i + 1, idxStatus + 1).setValue("Cancelled by System");
          sheet.getRange(i + 1, idxPaymentStatus + 1).setValue("Payment Denied");
        }
      }
    }
}

function createMidtransPaymentLink(appSheetPayload) {
  const orderId = appSheetPayload.orderId;
  const amount = appSheetPayload.amount;
  const guestEmail = appSheetPayload.guestEmail;
  const guestFirstName = appSheetPayload.guestFirstName;
  const apartmentName = appSheetPayload.apartmentName;

  const midtransPayload = {
      transaction_details: {
        order_id: orderId,
        gross_amount: amount
      },
      item_details: {
        price: amount,
        quantity: 1,
        name: apartmentName
      },
      credit_card: {
        secure: true,
        installment: {
          required: false,
          terms: {
            bni: [3, 6, 12],
            mandiri: [3, 6, 12],
            bca: [3, 6, 12],
            bri: [3, 6, 12],
            mega: [3, 6, 12]
          }
        }
      },
      customer_details: {
        first_name: guestFirstName,
        email: guestEmail
      },
      usage_limit: 5,
      expiry: {
        duration: 15,
        unit: "minutes"
      },
      callbacks: {
        finish: "https://www.appsheet.com/start/8b997da0-02f7-42cb-ad3f-b4621ab01292#view=BookingListGuest",
        error: "https://www.appsheet.com/start/8b997da0-02f7-42cb-ad3f-b4621ab01292#view=BookingListGuest"
      }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(midtransPayload),
    headers: {
      'Accept': 'application/json',
      'Authorization': 'Basic ' + Utilities.base64Encode(MIDTRANS_SERVER_KEY + ':')
    }
  };

  const midtransResponse = UrlFetchApp.fetch(MIDTRANS_SNAP_API_URL, options);
  const statusCode = midtransResponse.getResponseCode();
  const responseBody = midtransResponse.getContentText();
  const responseJson = JSON.parse(responseBody);
  
  if (statusCode === 201) {
    updatePaymentUrlToSheet(orderId, responseJson.redirect_url);
    MailApp.sendEmail({
      to: guestEmail,
      subject: "Payment Link for Your Apartment Booking",
      htmlBody: `
        <p>Hello ${guestFirstName},</p>
        <p>Your apartment booking request has been approved. Please continue by completing your payment using the following link:</p>
        <p><a href="${responseJson.redirect_url}">${responseJson.redirect_url}</a></p>
        <p>Amount due: <strong>Rp${Number(amount).toLocaleString("id-ID")}</strong></p>
        <br>
        <p>This email is automatically generated, please do not reply.</p>
      `
    });
  }
}

function updatePaymentUrlToSheet(orderId, paymentUrl) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const orderIdIndex = headers.indexOf('BookingID');
  const paymentUrlIndex = headers.indexOf('PaymentURL');

  for (let i = 1; i < data.length; i++) {
    if (data[i][orderIdIndex] === orderId) {
      sheet.getRange(i + 1, paymentUrlIndex + 1).setValue(paymentUrl);
    }
  }
}