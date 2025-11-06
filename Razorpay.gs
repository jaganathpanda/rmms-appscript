// Function to create an order on Razorpay
function createRazorpayOrder(paidAmount, paymentDate) {
  var razorpayKeyId = 'rzp_test_cXfTaJk9JKQ5Ff';  
  var razorpayKeySecret = 'xQevakKWjODHFnnSw1WblLPZ'; 
 var url = "https://api.razorpay.com/v1/orders";

  var orderData = {
    amount: Number(paidAmount) * 100, // Convert to paise
    currency: "INR",
    receipt: "receipt#" + new Date().getTime(),
    payment_capture: 1
  };

 var credentials = Utilities.base64Encode(razorpayKeyId + ":" + razorpayKeySecret);

  var options = {
    method: "post",
    headers: {
      Authorization: "Basic " + credentials,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(orderData),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());

  if (response.getResponseCode() === 200 || response.getResponseCode() === 201) {
    return {
      status: "Success",
      order_id: result.id,
      amount: result.amount,
      currency: result.currency
    };
  } else {
    return {
      status: "Error",
      message: result.error ? result.error.description : "Unknown error"
    };
  }
  return responseJson;
}
