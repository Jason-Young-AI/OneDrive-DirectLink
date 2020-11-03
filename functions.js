function deal_error(error){
  LogPanel.innerHTML = "Error !o!";
  OutputContent.value = error.toString();
}

function deal_cancel(){
  LogPanel.innerHTML = "Cancel >.<";
}

function deal_success(items){
  returned_items = items;
  console.log("%d items returned, Here is items data: %o", returned_items.value.length, returned_items);
  generate_direct_links(returned_items);
  LogPanel.innerHTML = "Success!";
}

function generate_direct_links(returned_items){
  var returned_items_array = returned_items.value;
  
  OutputContent.innerHTML = returned_items_array.length + "Items Selected.\n"
  if (returned_items_array.some(
    function(item){
      return item.shared == undefined || item.shared.scope != "anonymous";
    }
  )){
  	OutputContent.innerHTML += "There are items not shared, check the items that you selected.";
  }

  var direct_links_array = returned_items_array.map(
    function(item, index){
      var direct_link = get_direct_link(item, index);
      return direct_link;
    }
  );
  OutputContent.value = direct_links_array.join("\n");
}

function get_direct_link(item, index)
{
  var direct_link = eval("`http://storage.live.com/items/${item.id}:/${item.name}`");
  return direct_link;
}

var returned_items;

window.onload = function(){
  LogPanel = document.querySelector(".LogPanel");
  OutputContent = document.querySelector(".OutputContent");

  if (location.protocol != "https:"){
    var redirect = confirm("Now in HTTP mode! Function only take effect in HTTPS mode, Redirect(Y/N)?");
    if(redirect){
      location.protocol = "https:";
    }
  }
}

function LaunchOneDriveItemsPicker(){
  LogPanel.innerHTML = "Waiting for data to return ...";
  var odOptions = {
    clientId: "a85332cc-1dff-45ff-91ba-b183b844242d",
    action: "query",
    multiSelect: true,
    viewType: "all",
    openInNewWindow: true,
    advanced: {
      queryParameters: "select=audio,content,createdBy,createdDateTime,cTag,deleted,description,eTag,file,fileSystemInfo,folder,id,image,lastModifiedBy,lastModifiedDateTime,location,malware,name,package,parentReference,photo,publication,remoteItem,root,searchResult,shared,sharepointIds,size,specialFolder,video,webDavUrl,webUrl,activities,children,listItem,permissions,thumbnails,versions,@microsoft.graph.conflictBehavior,@microsoft.graph.downloadUrl,@microsoft.graph.sourceUrl"
    },
    success: function(files) {deal_success(files);},
    cancel: function() {deal_cancel();},
    error: function(error) {deal_error(error);}
  };
  OneDrive.open(odOptions);
}
