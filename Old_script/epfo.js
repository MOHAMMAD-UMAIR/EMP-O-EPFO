const http = require('https')

const websiteUrl = 'https://www.google.com';

// Make an HTTP GET request to the website
http.get(websiteUrl, (response) => {
    let data = '';
  
    // Accumulate the response data
    response.on('data', (chunk) => {
      data += chunk;
    });
  
    // Handle the completion of the response
    response.on('end', () => {
      console.log(data); // Output the response data
    });
  }).on('error', (error) => {
    console.error(`Error occurred: ${error.message}`);
  });