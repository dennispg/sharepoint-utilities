/// <reference path="./index.d.ts" />
SP.SOD.import(['sp.js'])
.then(() => {
    var context = new SP.ClientContext();
    var web = context.get_web();
    var lists = web.get_lists();
    var list = lists.getByTitle('Documents');
    var items: SP.ListItemCollection = list.get_queryResult("<View></View>");
    context.load(items);
    context.executeQuery()
    .then(() => {
        items.each((i, item) => { item.get_item('Title'); });
        items.every((item, i, items) => item.get_item('Title') == "") == false;
        items.find((item, i, items) => item.get_item('Title') == "").get_item('Title');
        var filtered1: SP.ListItem[] = items.filter(item => item.get_item('Title') == "");
        var filtered2: SP.ListItem[] = items.filter({"Title": "", "Author": web.get_currentUser()});
        items.firstOrDefault().get_item('Title');
        items.map((item, i) => item.get_item('Title')).length;
        items.some((item, i, items) => item.get_item('Title') == "") == false;
        items.toArray()[0].get_item('Title');
        var groups : {[group:string]:SP.ListItem[]} = items.groupBy(function(item) {
            return item.get_item('ContentType');
        });
        var total_sum: number = items.reduce(function(sum, item) {
            return sum + item.get_item('Total');
        }, 0);
    })
    var guid = SP.Guid.generateGuid();
});