/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("id_container").style.display = "none";
    document.getElementById("switch").addEventListener("change", function (event) {
      if (event.target.checked) {
        document.getElementById("strONOFF").innerText = "ON";
        document.getElementById("strONOFF").style.color = "blue";
        document.getElementById("id_container").style.display = "block";
      } else {
        document.getElementById("strONOFF").innerText = "OFF";
        document.getElementById("strONOFF").style.color = "red";
        document.getElementById("id_container").style.display = "none";
      }
    });
  }

  document.getElementById("btnInsert").onclick = insertSignature;
  document.getElementById("btnReset").onclick = resetAllField;
  document.getElementById("btnCallAPI").onclick = callSampleAPI;
});

function resetAllField() {
  document.getElementById("display_name").value = "";
  document.getElementById("email_id").value = "";
  document.getElementById("job_title").value = "";
  document.getElementById("phone_number").value = "";
  document.getElementById("greeting_text").value = "";
}

export function insertSignature() {
  // Get a reference to the current message
  //const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  //document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;

  let name = document.getElementById("display_name").value;
  let email = document.getElementById("email_id").value;
  let job_title = document.getElementById("job_title").value;
  let phone = document.getElementById("phone_number").value;
  let greeting = document.getElementById("greeting_text").value;

  let str = "";
  if (greeting.length > 0) {
    str += greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  str += "<td style='padding-left: 0px;'>";
  str += "<strong>" + name + "</strong>";
  str += "<br/>";
  str += job_title.length > 0 ? job_title + "<br/>" : "";
  str += email + "<br/>";
  str += phone.length > 0 ? phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  Office.context.mailbox.item.body.setSignatureAsync(
    str,
    {
      coercionType: "html",
    },

    function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
      }
    }
  );

  // Office.context.mailbox.item.body.setAsync(
  //   str,
  //   {
  //     coercionType: "html", // Write text as HTML
  //   },

  //   // Callback method to check that setAsync succeeded
  //   function (asyncResult) {
  //     if (asyncResult.status == Office.AsyncResultStatus.Failed) {
  //       write(asyncResult.error.message);
  //     }
  //   }
  // );
}

function callSampleAPI() {
  let json_str = "";
  fetch('https://randomuser.me/api/')
  .then(response => response.json())
  .then(json => 
    {
      json_str = json; 
      console.log(json_str.results[0].picture.large);

      document.getElementById("full_name").value = json_str.results[0].name.first + " " + json_str.results[0].name.last;
      document.getElementById("gender").value = json_str.results[0].gender;
      const date = new Date(json_str.results[0].dob.date);
      document.getElementById("dateofbirth").value = date.getMonth() + " - " + date.getDay() + " - " + date.getFullYear();
      document.getElementById("phone").value = json_str.results[0].phone;
      document.getElementById("avatar").src = json_str.results[0].picture.large;
    });
}
