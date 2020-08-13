var PartnerCode = "932030251";
var PartnerMainWarehouseCode = "Y30251_001";
var staff = ["031158", "031432", "032053", "032074", "032122", "031580", "032050", "031345"];
var onlineShop = "032058";


var today = new Date();
var day = today.getDate();
var month = today.getMonth();
var year = today.getFullYear();
var yesterday = new Date(Date.now() - 864e5);
var yesterdayDay = (yesterday.getDate() > 9 ? yesterday.getDate() : "0" + yesterday.getDate()).toString();
var yesterdayMonth = ((yesterday.getMonth()) > 9 ? yesterday.getMonth() + 1 : "0" + (yesterday.getMonth() + 1)).toString();
var yesterdayYear = yesterday.getFullYear().toString();
var yesterday2 = yesterdayYear + "" + yesterdayMonth + "" + yesterdayDay;
var InterfaceType = "ST";
var MessageName = "INVOICE";
var GUID = uuidv4();

// var DataPackagePeriod = "[DAILY, MONTHLY]";
var StatisticsPeriod = yesterdayYear + yesterdayMonth;
var ReportDate = today.toISOString().slice(0,10).replace(/-/g,"");
var ProcessStatus = "TRUE";

function uuidv4() {
	return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
		var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
		return v.toString(16);
	});
}
