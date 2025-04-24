document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("awsCredsForm").addEventListener("submit", async (e) => {
    e.preventDefault();

    // Get the IAM credentials
    const accessKey = document.getElementById("accessKey").value;
    const secretKey = document.getElementById("secretKey").value;

    // Simple validation
    if (!accessKey || !secretKey) {
      document.getElementById("output").textContent = "Error: Both Access Key and Secret Key are required!";
      return;
    }

    // Update AWS SDK config with user credentials
    AWS.config.update({
      region: "us-west-2", // Make sure this matches the region of your Lambda function
      credentials: new AWS.Credentials(accessKey, secretKey)
    });

    const lambda = new AWS.Lambda();

    // Display loading message
    document.getElementById("output").textContent = "Triggering Lambda function...";

    try {
      const result = await lambda.invoke({
        FunctionName: "CR_Desc",  // Ensure the function name matches your Lambda function
        Payload: JSON.stringify({
          subject: Office.context.mailbox.item.subject,
          bodyPreview: Office.context.mailbox.item.body ? Office.context.mailbox.item.body : "No body available"
        }),
      }).promise();

      // Display result from Lambda
      document.getElementById("output").textContent = JSON.stringify(result, null, 2);
    } catch (err) {
      // Display any error encountered during Lambda invocation
      document.getElementById("output").textContent = "Error: " + err.message;
    }
  });
});
