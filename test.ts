SP.SOD.import(['sp.js'])
.then(() => {
    var context = new SP.ClientContext();
    var lists = context.get_web().get_lists();
    var list = lists.getByTitle('Documents');
    var items = list.get_queryResult("<View></View>");
    context.load(items);
    context.executeQuery()
    .then(() => {
        items.each((i, item) => { item.get_item('Title'); });
        items.every((item, i, items) => item.get_item('Title') == "") == false;
        items.find((item, i, items) => item.get_item('Title') == "").get_item('Title');
        items.firstOrDefault().get_item('Title');
        items.map((item, i) => item.get_item('Title')).length;
        items.some((item, i, items) => item.get_item('Title') == "") == false;
        items.toArray()[0].get_item('Title');
    })
    var guid = SP.Guid.generateGuid();
});