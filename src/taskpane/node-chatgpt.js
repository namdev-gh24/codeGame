// // Replace YOUR_API_KEY with your actual API key
// const apiKey = "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";

// // Replace YOUR_MESSAGE with the message you want to send to the ChatGPT API
// const message = "Give random 10 vehicle names";

// // Set the API endpoint URL
// const apiUrl = "https://api.openai.com/v1/engines/davinci-codex/completions";

// // Set the request headers
// const headers = {
//   "Content-Type": "application/json",
//   Authorization: `Bearer ${apiKey}`,
// };

// // Set the request body
// const body = {
//   prompt: message,
//   max_tokens: 60,
//   n: 1,
//   stop: "\n",
// };

// // Send the HTTP request with jQuery
// $.ajax({
//   url: apiUrl,
//   type: "POST",
//   headers: headers,
//   data: JSON.stringify(body),
//   success: function (response) {
//     // Log the response from the ChatGPT API
//     console.log(response.choices[0].text);
//   },
//   error: function (xhr, status, error) {
//     // Log an error message if the request failed
//     console.error("Request failed. Returned status of " + xhr.status);
//   },
// });
