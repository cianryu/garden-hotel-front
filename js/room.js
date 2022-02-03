const current = new Date();
let year = current.getFullYear();
let month = current.getMonth()+1;
let day = current.getDate();
let fileCnt = 0;
let notCleaning = 0;
let vCnt = 0;
let cCnt = 0;
let oCnt = 0;
let ooCnt = 0;
let totalCnt = 0;
let sCnt = 0;

const a_type = ["3A"
              , "4A"
              , "5A"
              , "6A"
              , "7A"
              , "8A"
              , "9A"
              , "10A"
              , "11A"
              , "12A"
              , "14A"
              , "15A"
              , "16A"]
let a_staff = []

const b_type = ["4B"
              , "5B"
              , "6B"
              , "7B"
              , "8B"
              , "9B"
              , "10B"
              , "11B"
              , "12B"
              , "14B"
              , "15B"]
let b_staff = []

const c_type = ["4C"
              , "5C"
              , "6C"
              , "7C"
              , "8C"
              , "9C"
              , "10C"
              , "11C"
              , "12C"
              , "14C"
              , "15C"]
let c_staff = []

let roomTypeList = ["V"
                  , "C"
                  , "O"
                  , "O.O"
                  , "ⓥ"
]


window.addEventListener("load", function(event) {
  if(month < 10){
    month = "0" + month;
  }
  if(day < 10){
    day = "0" + day;
  }
  let currentDt = year + "년 " + month + "월 " + day + "일";
  document.getElementById("currentDt").innerHTML = currentDt;
  
  let roomTypeAll = document.getElementsByClassName("roomType");
  for(var i = 0 ; i < roomTypeAll.length ; i++){
    roomTypeAll[i].addEventListener("click", fn_change_room_type, false);
  }
});


function readExcel1() {
  if(sCnt < 1){
    alert("Summary를 업로드 후 진행해주시지 바랍니다.");
    document.getElementById("uploadBtn1").value = "";
    return false;
  }
  
  let dCurrentDt = month + "/" + day + "/" + year
  let input = event.target;
  let reader = new FileReader();
  reader.onload = function () {
    let data = reader.result;
    let workBook = XLSX.read(data, { type: 'binary' });
    console.log("sheetName : " + workBook.SheetNames[1]);
    workBook.SheetNames.forEach(function (sheetName) {
      if(sheetName == "Expected Departure List - Group"){
        return false;
      }
      let rows = XLSX.utils.sheet_to_json(workBook.Sheets[sheetName]);
      const datas = rows.map(parent => {
        return Object.keys(parent).reduce((acc, key) => ({
          ...acc,
          [key.replace(/\s/g, "")]: parent[key],
        }), {});
      });
      if(sheetName.split(" ")[1] != "Departure" && datas.length > 1){
        alert("Departure 문서가 아닙니다.");
        document.getElementById("uploadBtn1").value = "";
        return false;
      }
      if(datas[0].DepDate != dCurrentDt && datas.length > 1){
        alert("오늘 일자 Departure 문서가 아닙니다.");
        document.getElementById("uploadBtn1").value = "";
        return false;
      }
      var testNo = 0;
      datas.forEach(row => {
        let depDate = row.DepDate.replace(" ","");
        if(depDate != ""){
          var roomSId = document.getElementById("s_"+row.RmNo);
          if(roomSId != null && (roomSId.innerHTML != "C")){
            if(roomSId.innerHTML == "ⓥ"){
              console.log(testNo++);
            }
            roomSId.innerHTML = "C";
            roomSId.style.color = "red";
            ++cCnt;
            if(roomSId.innerHTML == "V"){
              --vCnt;
            }else if(roomSId.innerHTML == "O.O"){
              --ooCnt;
            }
          }
        }
      });
      fn_totalCnt();
    });
  };
  ++sCnt;
  reader.readAsBinaryString(input.files[0]);
}

function readExcel2() {
  let input = event.target;
  let reader = new FileReader();
  let sCurrentDt = year + "" + month + "" + day;
  console.log(sCurrentDt);
  reader.onload = function () {
    let data = reader.result;
    let workBook = XLSX.read(data, { type: 'binary' });
    workBook.SheetNames.forEach(function (sheetName) {
      let sheetVal = document.getElementById("uploadBtn2").value;
      console.log(sheetVal.split("_")[1].split(".")[0]);
      if(sheetName.split(" ")[2] != "Summary"){
        alert("Summary 문서가 아닙니다.");
        document.getElementById("uploadBtn2").value = "";
        return false;
      }else if(sheetVal.split("_")[1].split(".")[0] != sCurrentDt){
        alert("오늘 일자 Summary 문서가 아닙니다.");
        document.getElementById("uploadBtn2").value = "";
        return false;
      }
      
      let rows = XLSX.utils.sheet_to_json(workBook.Sheets[sheetName]);
      const datas = rows.map(parent => {
        return Object.keys(parent).reduce((acc, key) => ({
          ...acc,
          [key.replace(/\s/g, "")]: parent[key],
        }), {});
      });

      datas.forEach(row => {
        let roomStatus = "";
        let color = "red";
        let size = "medium";
        if((row.RoomStatus == "Vacant")){
          if(row.CleanStatus == "Cleaned"){
            roomStatus = "V";
            color = "black";
            size = "small";
          }else if(row.CleanStatus == "Dirty"){
            roomStatus = "ⓥ";
            size = "large";
          }
          delete row.RoomStatus;
          delete row.CleanStatus;
          for(key in row){
            var roomNo = row[key].split(" ")[0];
            var roomSId = document.getElementById("s_"+roomNo);
            if(roomSId != null && roomSId.innerHTML == ""){
              roomSId.innerHTML = roomStatus;
              roomSId.style.color=color;
              roomSId.style.fontSize=size;
              if(roomStatus == "V"){
                ++vCnt;
              }else if(roomStatus == "C" || roomStatus == "ⓥ"){
                ++cCnt;
              }
            }
          }
        }else if(row.RoomStatus == "Out Of Order"){
          roomStatus = "O.O"
          size="small";
          delete row.RoomStatus;
          delete row.CleanStatus;
          for(key in row){
            var roomNo = row[key].split(" ")[0];
            var roomSId = document.getElementById("s_"+roomNo);
            if(roomSId != null){
              roomSId.innerHTML = roomStatus;
              roomSId.style.color=color;
              roomSId.style.fontSize=size;
              ++ooCnt;
            }
          }
        }
        fn_totalCnt();
      });
    });
  };
  ++sCnt;
  reader.readAsBinaryString(input.files[0]);
  
}

function reRoomCheck(){
  if(sCnt == 0){
    alert("Summary를 업로드 후 진행해주시지 바랍니다.");
    return;
  }else if(sCnt == 1){
    alert("Departure를 업로드 후 진행해주시지 바랍니다.");
    return;
  }
  let startRoomNo = 1;
  let endRoomNo = 31;
  let floorNo = "";
  let roomNo = "";
  for(var i = 3 ; i <= 16 ; i++){
    if(i == 13){
      continue;
    }
    switch (i) {
      case 3 : 
        startRoomNo = 5;
        endRoomNo = 22;
        break;
      case 4 : 
        startRoomNo = 1;
        endRoomNo = 34;
        break;
      case 5 : 
        startRoomNo = 1;
        endRoomNo = 35;
        break;
      case 16 : 
        endRoomNo = 6;
        break;
    default:
      startRoomNo = 1;
      endRoomNo = 31;
      break;
    }
    if(i < 10){
      floorNo = "0" + i;
    }else{
      floorNo = i;
    }
    let roomStatus = "O";
    for(var j = startRoomNo ; j <= endRoomNo ; j++){
      if(j < 10){
        roomNo = floorNo + "0" + j;
      }else{
        roomNo = floorNo + "" + j;
      }
      var roomSId = document.getElementById("s_"+ roomNo);
      if(roomSId.innerHTML == "") {
        roomSId.innerHTML = roomStatus;
        ++oCnt;
      }
    }
  }
  ++sCnt
  fn_totalCnt();
}

function fn_totalCnt(){
  totalCnt = vCnt + cCnt + oCnt + ooCnt;
  document.getElementById("total_v").innerHTML = vCnt;
  document.getElementById("total_c").innerHTML = cCnt;
  document.getElementById("total_o").innerHTML = oCnt;
  document.getElementById("total_oo").innerHTML = ooCnt;
  document.getElementById("total").innerHTML = totalCnt;
}

function printPage(){
  printSetData();
  
  var initBody = document.body.innerHTML;
  window.onbeforeprint = function(){
    document.body.innerHTML = document.getElementById('content').innerHTML;
  }
  window.onafterprint = function(){
    document.body.innerHTML = initBody;
  }
  window.print();
  totalRoomChk("update");
  fn_floor_staff_re_set();
}

function fn_floor_staff(){
  
  if(sCnt < 3){
    alert("Occupied 처리 후 진행해 주시기 바랍니다.");
    return;
  }
  
  for(let i in a_type){
    var aStaff = document.getElementById(a_type[i]).value.replaceAll(" ","");
    document.getElementById(a_type[i]).value = aStaff;
    if(aStaff != undefined){
      var aRStaff = document.getElementsByClassName("staff"+a_type[i]);
      let aRoomNoClass = "floor" + a_type[i];
      let aRoomNo = document.getElementsByClassName(aRoomNoClass);
      for(var j = 0 ; j < aRStaff.length ; j++){
        if(aStaff == ""){
          aRoomNo[j].style.backgroundColor = "#fbffad";
        }else{
          aRoomNo[j].style.backgroundColor = "";
          aRStaff[j].value = aStaff;
        }
      }
      a_staff[i] = aStaff;
    }
  }
  for(i in b_type){
    var bStaff = document.getElementById(b_type[i]).value.replaceAll(" ","");
    document.getElementById(b_type[i]).value = bStaff;
    if(bStaff != undefined){
      var bRStaff = document.getElementsByClassName("staff"+b_type[i]);
      let bRoomNoClass = "floor" + b_type[i];
      let bRoomNo = document.getElementsByClassName(bRoomNoClass);
      for(var j = 0 ; j < bRStaff.length ; j++){
        if(bStaff == ""){
          bRoomNo[j].style.backgroundColor = "#fbffad";
        }else{
          bRoomNo[j].style.backgroundColor = "";
          bRStaff[j].value = bStaff;
        }
      }
      b_staff[i] = bStaff;
    }
  }
  for(i in c_type){
    var cStaff = document.getElementById(c_type[i]).value;
    if(cStaff != undefined){
      var cRStaff = document.getElementsByClassName("staff"+c_type[i]);
      let cRoomNoClass = "floor" + c_type[i];
      let cRoomNo = document.getElementsByClassName(cRoomNoClass);
      for(var j = 0 ; j < cRStaff.length ; j++){
        if(cStaff == ""){
          cRoomNo[j].style.backgroundColor = "#fbffad";
        }else{
          cRoomNo[j].style.backgroundColor = "";
          cRStaff[j].value = cStaff;
        }
      }
      b_staff[i] = bStaff;
    }
  }
  if(totalCnt == 372){
    fn_notCleaning();
  }
}

function fn_floor_staff_re_set(){
  for(var i = 0 ; i < 13 ; i++){
    if(a_staff[i] != null && a_staff[i] != undefined){
      document.getElementById(a_type[i]).value = a_staff[i];
    }
  }
  for(var i = 0 ; i < 11 ; i++){
    if(b_staff[i] != null && b_staff[i] != undefined){
      document.getElementById(b_type[i]).value = b_staff[i];
    }
  }
  for(var i = 0 ; i < 11 ; i++){
    if(c_staff[i] != null && c_staff[i] != undefined){
      document.getElementById(c_type[i]).value = c_staff[i];
    }
  }
}

function printSetData(){
  totalRoomChk("print");
}

function totalRoomChk(type){
  let startRoomNo = 1;
  let endRoomNo = 31;
  let floorNo = "";
  let roomNo = "";
  let input_html = "";

  fn_set_report(type);

  for(var i = 3 ; i <= 16 ; i++){
    if(i == 13){
      continue;
    }
    switch (i) {
      case 3 : 
        startRoomNo = 5;
        endRoomNo = 22;
        break;
      case 4 : 
        startRoomNo = 1;
        endRoomNo = 34;
        break;
      case 5 : 
        startRoomNo = 1;
        endRoomNo = 35;
        break;
      case 16 : 
        endRoomNo = 6;
        break;
    default:
      startRoomNo = 1;
      endRoomNo = 31;
      break;
    }
    if(i < 10){
      floorNo = "0" + i;
    }else{
      floorNo = i;
    }
    for(var j = startRoomNo ; j <= endRoomNo ; j++){
      if(j < 10){
        roomNo = floorNo + "0" + j;
      }else{
        roomNo = floorNo + "" + j;
      }

      var roomSId_input = document.getElementById("u_"+ roomNo).firstElementChild;
      var roomSId = document.getElementById("u_"+ roomNo);

      if(type == "update" && roomSId != null && roomSId != "") {
        input_html = fn_set_input(i, j);
        
        var roomSId_text = roomSId.innerHTML;
        roomSId.innerHTML = input_html;
        roomSId.firstElementChild.value = roomSId_text;
      }else if(type == "print" && roomSId_input != null && roomSId_input != "") {
        var roomVal = roomSId_input.value.replaceAll(" ", "");
        roomSId_input.parentElement.innerHTML = roomVal;
      }
    }
  }
}

function fn_set_report(type){
  let nightStaff_input = document.getElementById("nightStaff").firstElementChild;
  let nightStaff = document.getElementById("nightStaff");
  let staffNum_input = document.getElementById("staffNum").firstElementChild;
  let staffNum = document.getElementById("staffNum");

  if(type == "update") {
    input_html = fn_set_input(0, 0);
    if(nightStaff != null && nightStaff != ""){
      let text = nightStaff.innerHTML;
      nightStaff.innerHTML = input_html;
      nightStaff.firstElementChild.value = text;
    }
    if(staffNum != null && staffNum != ""){
      text = staffNum.innerHTML;
      staffNum.innerHTML = input_html;
      staffNum.firstElementChild.value = text;
    }
  }else if(type == "print") {
    if(nightStaff_input != null && nightStaff_input != ""){
      let value = nightStaff_input.value.replaceAll(" ", "");
      nightStaff_input.parentElement.innerHTML = value;
    }
    if(staffNum_input != null && staffNum_input != ""){
      let value = staffNum_input.value.replaceAll(" ", "");
      staffNum_input.parentElement.innerHTML = value;
    }
  }
}

function fn_set_input(i, j){
  switch (i) {
    case 0 :
      input_html = '<input type="text" value=""/>';
      break;
    case 3 : 
      input_html = '<input type="text" class="staff'+i+'A" value=""/>'
      break;
    case 4 : 
      if((j > 5 && j < 11) || (j > 21 && j < 28)){
        input_html = '<input type="text" class="staff'+i+'A" value=""/>'
      }else if(j > 10 && j < 22){
        input_html = '<input type="text" class="staff'+i+'B" value=""/>'
      }else{
        input_html = '<input type="text" value=""/>';
      }
      break;
    case 5 : 
      if((j > 5 && j < 11) || (j > 22 && j < 29)){
        input_html = '<input type="text" class="staff'+i+'A" value=""/>'
      }else if(j > 10 && j < 23){
        input_html = '<input type="text" class="staff'+i+'B" value=""/>'
      }else{
        input_html = '<input type="text" value=""/>';
      }
      break;
    case 16 : 
      input_html = '<input type="text" class="staff'+i+'A" value=""/>'
      break;
  default:
    if((j > 5 && j < 11) || (j > 21 && j < 27)){
      input_html = '<input type="text" class="staff'+i+'A" value=""/>'
    }else if(j > 10 && j < 22){
      input_html = '<input type="text" class="staff'+i+'B" value=""/>'
    }else{
      input_html = '<input type="text" value=""/>';
    }
    break;
  }
  return input_html;
}

function fn_change_room_type(){
  let roomType = this.innerHTML;
  let nextRoomType;
  switch (roomType) {
    case roomTypeList[0] : 
      this.style.color = "red";
      this.style.fontSize = "medium";
      nextRoomType = roomTypeList[1];
      vCnt--;
      cCnt++;
      break;
    case roomTypeList[1] : 
      this.style.color = "black";
      nextRoomType = roomTypeList[2];
      cCnt--;
      oCnt++;
      break;
    case roomTypeList[2] :
      this.style.color = "red";
      this.style.fontSize = "small";
      nextRoomType = roomTypeList[3];
      oCnt--;
      ooCnt++;
      break;
    case roomTypeList[3] : 
      this.style.color = "red";
      this.style.fontSize = "large";
      nextRoomType = roomTypeList[4];
      ooCnt--;
      vCnt++;
      break;
    case roomTypeList[4] : 
      this.style.color = "black";
      this.style.fontSize = "small";
      nextRoomType = roomTypeList[0];
      ooCnt--;
      vCnt++;
      break;
    case "" : 
      this.style.color = "black";
      nextRoomType = roomTypeList[0];
      vCnt++;
      break;
  }
  this.innerHTML = nextRoomType;
  fn_totalCnt();
  if(totalCnt == 372){
    fn_notCleaning();
  }
}

function fn_notCleaning(){
  notCleaning = 0;
  if(sCnt < 3){
    alert("Occupied 처리 후 진행해 주시기 바랍니다.");
    return;
  }
  let startRoomNo = 1;
  let endRoomNo = 31;
  let floorNo = "";
  let roomNo = "";
  for(var i = 3 ; i <= 16 ; i++){
    if(i == 13){
      continue;
    }
    switch (i) {
      case 3 : 
        startRoomNo = 5;
        endRoomNo = 22;
        break;
      case 4 : 
        startRoomNo = 1;
        endRoomNo = 34;
        break;
      case 5 : 
        startRoomNo = 1;
        endRoomNo = 35;
        break;
      case 16 : 
        endRoomNo = 6;
        break;
    default:
      startRoomNo = 1;
      endRoomNo = 31;
      break;
    }
    if(i < 10){
      floorNo = "0" + i;
    }else{
      floorNo = i;
    }
    let roomStatus = "O";
    for(var j = startRoomNo ; j <= endRoomNo ; j++){
      if(j < 10){
        roomNo = floorNo + "0" + j;
      }else{
        roomNo = floorNo + "" + j;
      }
      var roomSId = document.getElementById("s_" + roomNo);
      var staffNm = document.getElementById("u_" + roomNo).firstElementChild.value;
      if((roomSId.innerHTML == "C" || roomSId.innerHTML == "ⓥ" || roomSId.innerHTML == "O") &&
          staffNm == "") {
        roomSId.style.backgroundColor = "#ffe3ff";
        notCleaning++;
      }else{
        roomSId.style.backgroundColor = "";
      }
    }
  }
  document.getElementById("notCleaningCnt").innerHTML = notCleaning;
}