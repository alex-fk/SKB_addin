/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
let item;
let OutlookUserMail;
// Office is ready.
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    item = Office.context.mailbox.item;
    // console.log(item);
		OutlookUserMail = Office.context.mailbox.userProfile.emailAddress;
    $('#input-user-mail').val(OutlookUserMail).change();

    // Заполнение места
    function setLocation(selected_room) {
      item.location.setAsync(
        selected_room.room_link_guest,
        function (asyncResult) {
          if (asyncResult.status == Office.AsyncResultStatus.Failed){
            console.log("Место не установлено");
            console.log(asyncResult.error.message);
          } else {
            console.log("Место успешно установлено");
            // Successfully set the location.
            // Do whatever is appropriate for your scenario,
            // using the arguments var1 and var2 as applicable.
          }
        }
      )
    };

    // Gets the subject of the item that the user is composing.
    function getSubject() {
      return item.subject.getAsync((asyncResult) => {
           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
               console.log(asyncResult.error.message);
               return;
           }
             // Display the subject on the page.
           // console.log(asyncResult);
           return asyncResult.value;
       });
     };
    // Заполнение темы
    function setSubject() {
      const subject = `Новая встреча`;
      item.subject.setAsync(
        subject,
        { asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
          } else {
            console.log("Успешно установлено название встречи");
            /*
            The subject was successfully set.
            Run additional operations appropriate to your scenario and
            use the optionalVariable1 and optionalVariable2 values as needed.
            */
          }
        }
      )
    };

    // Вставка текста в сообщение
    function setItemBody(selected_room) {
      // Identify the body type of the mail item.
      item.body.getTypeAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.log(asyncResult.error.message);
              return;
          }

          // Insert data of the appropriate type into the body.
          if (asyncResult.value === Office.CoercionType.Html) {
              // Insert HTML into the body.
              item.body.setSelectedDataAsync(
                  "<b>Название комнаты:</b> "+selected_room.room_name+"<br><b> Подключение по ссылке: </b><a href = '"+selected_room.room_link_guest+"' title = 'Нажмите для подключения к встрече'>"+selected_room.room_link_guest+"</a><br><b> Подключение с ВКС терминала: </b>"+selected_room.room_sip_uri+"<br><b> Подключение по телефону: </b>"+selected_room.room_tel_number+"<br>",
                  { coercionType: Office.CoercionType.Html},
                  (asyncResult) => {
                      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                          console.log(asyncResult.error.message);
                          return;
                      }

                      /*
                        Run additional operations appropriate to your scenario and
                        use the optionalVariable1 and optionalVariable2 values as needed.
                      */
              });
          }
          else {
              // Insert plain text into the body.
              item.body.setSelectedDataAsync(
                  "Название комнаты: "+selected_room.room_name+ "\nПодключение по ссылке: "+selected_room.room_link_guest+"\nПодключение с ВКС терминала: "+selected_room.room_sip_uri+"\nПодключение по телефону: "+selected_room.room_tel_number+"\n",
                  { coercionType: Office.CoercionType.Text},
                  (asyncResult) => {
                      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                          console.log(asyncResult.error.message);
                          return;
                      }

                      /*
                        Run additional operations appropriate to your scenario and
                        use the optionalVariable1 and optionalVariable2 values as needed.
                      */
              });
          }
      });
    }
  }

  

// $().ready(function() {
    //Поиск комнат
    let find_rooms_arr = [];
    $('#input-search-room').autocomplete({
    //Запрос списка комнат JSON
      // classes: {
      //     "ui-menu-item": "icon-list-mts"
      // },
      minLength: 2,
      create: function() {
        $(this).data('ui-autocomplete')._renderItem = function( ul, item ) {
          if (item.room_vendor===1) {
            return $('<li>')
              .data("item.autocomplete", item)
              .append($("<div>").addClass('icon-list-scb').html(item.label))
              .appendTo(ul);      
          }
          else if (item.room_vendor===2) {
            return $('<li>')
              .data("item.autocomplete", item)
              .append($("<div>").addClass('icon-list-mts').html(item.label))
              .appendTo(ul);     
          }
        }
      },
      source: function(req, add){
        //Передаём запрос на сервер
        $.getJSON("https://localhost:8000/rooms/"+req.term, function(data) {
          //Создаем массив для объектов ответа
          let suggestions = [];
          //Обрабатываем ответ
          find_rooms_arr.length = 0;
          $.each(data, function(i, val){
            if (val.room_vendor===2) {
              let room_name_witd_id = val.room_name +" " +val.room_id;
              suggestions.push({room_id:val.room_id, value:room_name_witd_id, room_vendor:val.room_vendor, room_link_guest:val.room_link_guest, 
              room_link_host:val.room_link_host, label:room_name_witd_id, room_name:val.room_name, room_owner:val.room_owner, room_sip_uri:val.room_sip_uri,
              room_tel_number:val.room_tel_number});
              find_rooms_arr.push(val);
            }
            else{
              suggestions.push({room_id:val.room_id, value:val.room_name, room_vendor:val.room_vendor, room_link_guest:val.room_link_guest, 
              room_link_host:val.room_link_host, label:val.room_name, room_name:val.room_name, room_owner:val.room_owner, room_sip_uri:val.room_sip_uri,
              room_tel_number:val.room_tel_number});
              find_rooms_arr.push(val);              
            }

          });
          //Передаем массив обратному вызову
          add(suggestions);
        });
      },
    });


    
    $('#input-search-room').on("focus", function() {
		if ($('#input-search-room').hasClass("ui-state-error")){
			$('#input-search-room').removeClass("ui-state-error");
		}
    });

    let selected_room = null
    $('#input-search-room').on("autocompleteselect", function( event, ui ) {
      console.log('------селекс комплит') 
      console.log(ui);
      selected_room = ui.item;
      console.log(selected_room);
        // console.log('fdsfdsf')
        // console.log(find_rooms_arr);
        // console.log(ui.item.value);
        // selected_room = find_rooms_arr.find((room) => room.room_name === ui.item.value);
        // console.log(selected_room);
    });
    $('#input-search-room').on("autocompletechange", function( event, ui ) {
        console.log(selected_room); 
        console.log('------чандже'); 
        console.log(ui.item);

      if (ui !== null && ui.item !==null ) {
          console.log("свойство существует");
          // Проверяем, является ли элемент объектом
            if (ui.item.value != selected_room.value){
              selected_room = null; 
              alertInput('#input-search-room');
            }
      } else {
        console.log("Если объект не существует");
        selected_room = null; 
      }
       
          // console.log('fdsfdsf')
          // console.log(find_rooms_arr);
          // console.log(ui.item.value);
          // selected_room = find_rooms_arr.find((room) => room.room_name === ui.item.value);
          // console.log(selected_room);
          console.log(selected_room);
    });


	$("#form-search-room").on("submit", function( event ) {
		let sel_room = $('#input-search-room').val();
		console.log(sel_room);
		console.log(selected_room);
		
		if ((sel_room) && (selected_room != null)){
			console.log("Form Submitted");
			console.log(sel_room);

      item.subject.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
            return;
        }

        // Display the subject on the page.
        if (asyncResult.value == ""){
          setSubject();
        };
    });

			setLocation(selected_room);
			setItemBody(selected_room);
		} else {
			alertInput('#input-search-room');
		}	
		event.preventDefault();
  });

function alertInput( el ) {
	$(el).addClass( "ui-state-error" );
}


});
// });
