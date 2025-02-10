/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(function() {
  // Office is ready.
  $(document).ready(function () {
    // The document is ready.
    var OutlookUserMail = Office.context.mailbox.userProfile.emailAddress;
    $('#input-user-mail').val(OutlookUserMail).change();
  });
});
var rooms_arr = [];
$(document).ready(function () {
  // The document is ready.
   $('#input-user-mail').val('avedisovai@sovcombank.ru').change();

  //поиск комнат
  $('#input-search-room').autocomplete({
//Определяем обратный вызов к результатам форматирования
    minLength: 3,
    source: function(req, add){
      //Передаём запрос на сервер
      $.getJSON("http://localhost:8000/rooms/"+req.term, function(data) {
        //Создаем массив для объектов ответа
       
        var suggestions = [];

        //Обрабатываем ответ
        $.each(data, function(i, val){								
          suggestions.push(val.room_name);
        });
        //Передаем массив обратному вызову
        add(suggestions);
        
      });
    },


  });
  $('#input-search-room' ).on( "autocompletechange", function( event, ui ) {
    console.log('fdsfdsf')
    console.log(ui);
  } );

});

// Office.onReady(function(info) => {
//   if (info.host === Office.HostType.Outlook) {
//     // document.getElementById("sideload-msg").style.display = "none";
//     // document.getElementById("app-body").style.display = "flex";
//     // document.getElementById("run").onclick = run;
//     var OutlookUserMail = Office.context.mailbox.userProfile.emailAddress;
//     document.getElementById("input-user-mail").value(OutlookUserMail);
//     console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
//   });
// }


// export async function run() {
//   /**
//    * Insert your Outlook code here
//    */

//   const item = Office.context.mailbox.item;
//   let insertAt = document.getElementById("item-subject");
//   let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
//   insertAt.appendChild(label);
//   insertAt.appendChild(document.createElement("br"));
//   insertAt.appendChild(document.createTextNode(item.subject));
//   insertAt.appendChild(document.createElement("br"));
// };