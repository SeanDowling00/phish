Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  const item = Office.context.mailbox.item;
  var fraudInfo = "test";

  const axios = require('axios');

// Replace YOUR_API_KEY with your actual VirusTotal API key
const apiKey = '813861dc28625873a20aa3dddcd0b33471ab4bc15063a8268a422d199b75506c';

// The file hash you want to scan
const fileHash = '8fa2b03af4e1e5c101ab000a0e6fda6745359871dbb040895d95d689960491f0';

// Build the API request URL
const url = `https://www.virustotal.com/vtapi/v2/file/report?apikey=${apiKey}&resource=${fileHash}`;

// Send the API request and receive the response
axios.get(url)
  .then(response => {
    // Decode the response data
    const responseData = response.data;
    
    // Check the response for a scan result
    if (responseData.positives && responseData.positives > 0) {
      console.log('File is infected with a virus');
      
    document.getElementById("item-attachscan").innerHTML = "<b>Malicious</b>";

    } else {
      document.getElementById("item-attachscan").innerHTML = "<b>Clean</b>";
    }
  })
  .catch(error => {
    console.error(error);
    // Add an error message to the UI, to notify the user that something went wrong
    console.log("Error:", error);
    document.getElementById("item-attachscan").innerHTML = "An error occurred while scanning the file for viruses." + error;
    });


  Office.context.mailbox.item.getAllInternetHeadersAsync(processHeaders);
  
  function processHeaders(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      
      var attachmentNames = "";
      for (var i = 0; i < item.attachments.length; i++) {
        attachmentNames += item.attachments[i].name + ", ";
      }
  
      if (attachmentNames) {
        attachmentNames = attachmentNames.substring(0, attachmentNames.length - 2);
      } else {
        attachmentNames = "No attachments found";
      }
      // Extract SPF, DKIM, and DMARC headers
      var spf = asyncResult.value.match(/spf=([^;\s]+)/g);
      if (spf) {
        spf = spf[0].substring(4);
      } else {
        spf = "Not Found";
      }
  
      var dkim = asyncResult.value.match(/dkim=([^;\s]+)/g);
      if (dkim) {
        dkim = dkim[0].substring(5);
      } else {
        dkim = "Not Found";
      }
  
      var dmarc = asyncResult.value.match(/dmarc=([^;\s]+)/g);
      if (dmarc) {
        dmarc = dmarc[0].substring(6);
      } else {
        dmarc = "Not Found";
      }
  
      document.getElementById("item-results").classList.remove("fraud");
      document.getElementById("item-results").classList.remove("authentic");
  
      if (spf === "Not Found" || dkim === "Not Found" || dmarc === "Not Found" || spf === "none" 
          || dkim === "none" || dmarc === "none" || spf === "fail" 
          || dkim === "fail" || dmarc === "fail") 
        {
        var results = "Potentially Fraudulent";
        document.getElementById("item-results").classList.add("fraud");
        var fraudInfo = "This email may be fraudulant, this is either because <br/> A) The analysis returned failed results <br/> B) The checks could not be complete";
        document.getElementById("item-fraudInfo").classList.add("fraud");
        var proceed = "This analysis does not mean this email is 100% fraudulant or authentic, spoofing attacks can still occur!!";
        document.getElementById("item-proceed").classList.add("fraud");
      }
    else 
      {
        var results = "Authentic Email";
        document.getElementById("item-results").classList.add("authentic");
        var fraudInfo = "This is more than likely an authentic Email, this means the email may <br/> A) Not be Spoofed <br/> B) Not be a Phishing Email";
        document.getElementById("item-fraudInfo").classList.add("authentic");
        var proceed = "This analysis does not mean this email is 100% fraudulant or authentic, spoofing attacks can still occur!";
        document.getElementById("item-proceed").classList.add("fraud");
      }
      var info = "Click each button below for more information";
      
    
      //===============================================================================================
      
        
       // Write message property value to the task pane
    document.getElementById("item-results").innerHTML = results;
    document.getElementById("item-fraudInfo").innerHTML = fraudInfo;
    document.getElementById("item-proceed").innerHTML = proceed;
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    document.getElementById("item-sender").innerHTML = "<b>Sender:</b> <br/>" + item.sender.emailAddress;
    document.getElementById("item-date").innerHTML = "<b>Received:</b> <br/>" + item.dateTimeCreated;
    document.getElementById("item-spf").innerHTML = "<b>SPF:</b> <br/>" + spf;
    document.getElementById("item-dmarc").innerHTML = "<b>DMARC:</b> <br/>" + dmarc;
    document.getElementById("item-dkim").innerHTML = "<b>DKIM:</b> <br/>" + dkim;
    document.getElementById("item-attachment").innerHTML = "<b>Attachment Name:</b> <br/>" + attachmentNames;
    document.getElementById("item-info").innerHTML = info;

  }

}}