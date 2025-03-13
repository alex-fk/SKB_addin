(function () {
  "use strict";

// Office is ready.
Office.onReady(function (info) {
  let item;
  let OutlookUserMail;



  if (info.host === Office.HostType.Outlook) {
    item = Office.context.mailbox.item;
		OutlookUserMail = Office.context.mailbox.userProfile.emailAddress;
    $('#input-user-mail').val(OutlookUserMail).change();

    // Заполнение места
    // function setLocation(selected_room) {
    //   Office.context.mailbox.item.location.setAsync(
    //     selected_room.room_link_guest,
    //     function (asyncResult) {
    //       if (asyncResult.status == Office.AsyncResultStatus.Failed){
    //       } else {
    //       }
    //     }
    //   )
    // };

    // function getSubject() {
    //   return Office.context.mailbox.item.subject.getAsync((asyncResult) => {
    //        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    //            return;
    //        }
    //        return asyncResult.value;
    //    });
    //  };
    //  Заполнение темы
    // function setSubject() {
    //   Office.context.mailbox.item.subject.setAsync(
    //     "Новая встреча",        
    //     (asyncResult) => {
    //       if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    //         return;
    //       }
    //     }
    //   )
    // };

    // Вставка текста в сообщение
    // function setItemBody(selected_room) {
    //   Office.context.mailbox.item.body.getTypeAsync((asyncResult) => {
    //       if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    //           return;
    //       }
    //       if (asyncResult.value === Office.CoercionType.Html) {
    //           // Insert HTML into the body.
    //           Office.context.mailbox.item.body.prependAsync(
    //               "<b>Название комнаты:</b> "+selected_room.room_name+"<br><b> Подключение по ссылке: </b><a href = '"+selected_room.room_link_guest+"' title = 'Нажмите для подключения к встрече'>"+selected_room.room_link_guest+"</a><br><b> Подключение с ВКС терминала: </b>"+selected_room.room_sip_uri+"<br><b> Подключение по телефону: </b>"+selected_room.room_tel_number+"<br>",
    //               { coercionType: Office.CoercionType.Html},
    //               (asyncResult) => {
    //                   if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    //                     console.log(asyncResult.error.message);
    //                     return;
    //                   }
    //           });
    //       }
    //       else {
    //           // Insert plain text into the body.
    //           Office.context.mailbox.item.body.prependAsync(
    //               "Название комнаты: "+selected_room.room_name+ "\nПодключение по ссылке: "+selected_room.room_link_guest+"\nПодключение с ВКС терминала: "+selected_room.room_sip_uri+"\nПодключение по телефону: "+selected_room.room_tel_number+"\n",
    //               { coercionType: Office.CoercionType.Text},
    //               (asyncResult) => {
    //                   if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    //                      console.log(asyncResult.error.message);
    //                      return;
    //                   }
    //           });
    //       }
    //   });
    // }
  }


  

    // Поиск комнат
    let find_rooms_arr = [];
    $('#input-search-room').autocomplete({
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
        $.getJSON("https://scbms.sovcombank.group/rooms/"+req.term, function(data) {
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
              room_tel_number:"8 800 707 1350 доб. " + val.room_tel_number});
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
      selected_room = ui.item;
        // selected_room = find_rooms_arr.find((room) => room.room_name === ui.item.value);
    });
    $('#input-search-room').on("autocompletechange", function( event, ui ) {
      if (ui !== null && ui.item !==null ) {
            if (ui.item.value != selected_room.value){
              selected_room = null; 
              alertInput('#input-search-room');
            }
      } else {
        selected_room = null; 
      }
    });


	$("#form-search-room").on("submit", function( event ) {
		let sel_room = $('#input-search-room').val();
		if ((sel_room) && (selected_room != null)){
        Office.context.mailbox.item.subject.getAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              write(asyncResult.error.message);
              return;
          }
          if (asyncResult.value == ""){
            Office.context.mailbox.item.subject.setAsync(
              "Новая встреча",        
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  return;
                }
              }
            )
          };
        });

        Office.context.mailbox.item.location.setAsync(
          selected_room.room_link_guest,
          function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
            } else {
            }
          }
        );
        let content_htm = `Название комнаты: ${selected_room.room_name}<br>Подключение по ссылке: <a href = "${selected_room.room_link_guest}">${selected_room.room_link_guest}</a><br>Подключение с ВКС терминала: ${selected_room.room_sip_uri}<br>Подключение по телефону: ${selected_room.room_tel_number}<br>`
        let content_text = `Название комнаты: ${selected_room.room_name}\nПодключение по ссылке: ${selected_room.room_link_guest}\nПодключение с ВКС терминала: ${selected_room.room_sip_uri}\nПодключение по телефону: ${selected_room.room_tel_number}\n`

        Office.context.mailbox.item.body.getTypeAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.log(asyncResult)
              return;
          }
          
          if (asyncResult.value === Office.CoercionType.Html) {
              // Insert HTML into the body.
              Office.context.mailbox.item.body.prependAsync(
                  content_htm,
                  { coercionType: Office.CoercionType.Html },
                  (asyncResult) => {
                      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                          console.log(selected_room);
                          console.log(asyncResult);
                          console.log(asyncResult.error.message);
                          // Insert plain text into the body.
                          Office.context.mailbox.item.body.prependAsync(
                            content_text, 
                            {coercionType: Office.CoercionType.Text},
                            (asyncResult) => {
                            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                console.log(asyncResult);
                                console.log(asyncResult.error.message);
                                return;
                            }
                            });
                            return;
                        }
              });
          }
          else {
              // Insert plain text into the body.
              Office.context.mailbox.item.body.prependAsync(
                  content_text,
                  {coercionType: Office.CoercionType.Text},
                  (asyncResult) => {
                      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                          console.log(asyncResult);
                          console.log(asyncResult.error.message);
                          return;
                      }
              });
          }
        });



   

		} else {
			alertInput('#input-search-room');
		}	
		event.preventDefault();
  });

  function alertInput( el ) {
	  $(el).addClass( "ui-state-error" );
  }

});
})();