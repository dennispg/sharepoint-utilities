class ClientObjectCollection {
    each(callback) {
        return this.forEach((item, i) => {
            callback(i, item);
        });
    }
    forEach(iteratee) {
        var index = 0, enumerator = this.getEnumerator();
        while (enumerator.moveNext()) {
            if (iteratee(enumerator.get_current(), index++, this) === false)
                break;
        }
    }
    map(iteratee) {
        var index = -1, enumerator = this.getEnumerator(), result = [];
        while (enumerator.moveNext()) {
            result[++index] = iteratee(enumerator.get_current(), index, this);
        }
        return result;
    }
    toArray() {
        var collection = [];
        this.each((i, item) => {
            collection.push(item);
        });
        return collection;
    }
    some(iteratee) {
        var val = false;
        this.each((i, item) => {
            if (iteratee(item, i, this)) {
                val = true;
                return false;
            }
        });
        return val;
    }
    every(iteratee) {
        var val = true;
        var hasitems = false;
        this.each((i, item) => {
            hasitems = true;
            if (!iteratee(item, i, this)) {
                val = false;
                return false;
            }
        });
        return hasitems && val;
    }
    find(iteratee) {
        var val = undefined;
        this.each((i, item) => {
            if (iteratee(item, i, this)) {
                val = item;
                return false;
            }
        });
        return val;
    }
    reduce(iteratee, accumulator) {
        this.forEach((item, i) => {
            accumulator = iteratee(accumulator, item, i, this);
        });
        return accumulator;
    }
    groupBy(iteratee) {
        return this.reduce((result, value) => {
            var key = iteratee(value);
            if (Object.prototype.hasOwnProperty.call(result, key))
                result[key].push(value);
            else
                result[key] = [value];
            return result;
        }, {});
    }
    matches(source) {
        return item => {
            for (var prop in source) {
                var compare_val = source[prop];
                var value = null;
                if (item instanceof SP.ListItem) {
                    value = item.get_item(prop);
                    if (value instanceof SP.FieldLookupValue || value instanceof SP.FieldUserValue) {
                        value = value.get_lookupId();
                        if (compare_val instanceof SP.User)
                            compare_val = compare_val.get_id();
                        else if (compare_val instanceof SP.FieldLookupValue || compare_val instanceof SP.FieldUserValue)
                            compare_val = compare_val.get_lookupId();
                    }
                    else if (value instanceof SP.FieldUrlValue) {
                        value = value.get_url() + ';#' + value.get_description();
                        if (compare_val instanceof SP.FieldUrlValue)
                            compare_val = compare_val.get_url() + ';#' + compare_val.get_description();
                    }
                }
                else
                    value = item[prop];
                return value == compare_val;
            }
        };
    }
    property(props) {
        var properties = Array.isArray(props) ? props : [props];
        return item => {
            if (item instanceof SP.ListItem)
                return properties.every(prop => item.get_item(prop));
            else
                return properties.every(prop => item[prop]);
        };
    }
    filter(predicate) {
        var predicateType = Object.prototype.toString.call(predicate);
        switch (predicateType) {
            case '[object Function]':
                var items = [];
                var filter = predicate;
                this.forEach((item, i) => {
                    if (filter(item, i, this))
                        items.push(item);
                });
                return items;
            case 'object':
                return this.filter(this.matches(predicate));
            case 'string':
            case '[object Array]':
                return this.filter(this.property(predicate));
        }
        return null;
    }
    firstOrDefault(iteratee) {
        var index = 0, enumerator = this.getEnumerator();
        if (enumerator.moveNext()) {
            var current = enumerator.get_current();
            if (iteratee) {
                if (iteratee(current, index++, this))
                    return current;
            }
            else
                return current;
        }
        return null;
    }
}
var rejectionHandler;
export function registerUnhandledErrorHandler(handler) {
    rejectionHandler = handler;
}
class ClientContext {
    executeQuery() {
        var context = this;
        return new Promise((resolve, reject) => {
            context.executeQueryAsync((sender, args) => { resolve(args); }, (sender, args) => { reject(args); });
        })
            .catch((args) => {
            if (rejectionHandler)
                return rejectionHandler(args);
            return args;
        });
    }
}
class List {
    get_queryResult(queryText) {
        var query = new SP.CamlQuery();
        query.set_viewXml(queryText || "<View><Query></Query></View>");
        return this.getItems(query);
    }
}
class UserCustomActionCollection {
    ensure(custom_action) {
        var actions = this;
        var action;
        var context = actions.get_context();
        context.load(actions);
        return context.executeQuery()
            .then(() => {
            action = actions.filter(a => a.get_name() == custom_action.name || a.get_title() == custom_action.title)[0];
            if (!action)
                action = actions.add();
            Object.getOwnPropertyNames(custom_action).forEach(prop => {
                switch (prop) {
                    case "commandUIExtension":
                        custom_action.commandUIExtension && action.set_commandUIExtension(custom_action.commandUIExtension);
                        break;
                    case "description":
                        custom_action.description && action.set_description(custom_action.description);
                        break;
                    case "group":
                        custom_action.group && action.set_group(custom_action.group);
                        break;
                    case "imageUrl":
                        custom_action.imageUrl && action.set_imageUrl(custom_action.imageUrl);
                        break;
                    case "location":
                        custom_action.location && action.set_location(custom_action.location);
                        break;
                    case "name":
                        custom_action.name && action.set_name(custom_action.name);
                        break;
                    case "registrationType":
                        custom_action.registrationType && action.set_registrationType(custom_action.registrationType);
                        break;
                    case "rights":
                        custom_action.rights && action.set_rights(custom_action.rights);
                        break;
                    case "scriptBlock":
                        custom_action.scriptBlock && action.set_scriptBlock(custom_action.scriptBlock);
                        break;
                    case "scriptSrc":
                        custom_action.scriptSrc && action.set_scriptSrc(custom_action.scriptSrc);
                        break;
                    case "sequence":
                        custom_action.sequence && action.set_sequence(custom_action.sequence);
                        break;
                    case "title":
                        custom_action.title && action.set_title(custom_action.title);
                        break;
                    case "url":
                        custom_action.url && action.set_url(custom_action.url);
                        break;
                }
            });
            action.update();
            context.load(action);
            return context.executeQuery();
        })
            .then(() => {
            return action.get_id();
        });
    }
    remove(selector) {
        var actions = this;
        var context = actions.get_context();
        context.load(actions);
        return context.executeQuery()
            .then(() => {
            var action = actions.filter(a => selector(a))[0];
            if (action)
                action.deleteObject();
            return context.executeQuery();
        });
    }
    removeByTitle(title) {
        return this.remove(action => action.get_title() == title);
    }
    removeByName(name) {
        return this.remove(action => action.get_name() == name);
    }
    removeById(id) {
        return this.remove(action => action.get_id().equals(id));
    }
}
class Guid {
    static generateGuid() {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
            var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }
    ;
}
var sodBaseAddress = null;
var getSodBaseAddress = function () {
    if (sodBaseAddress)
        return sodBaseAddress;
    if (_v_dictSod['sp.js'] && _v_dictSod['sp.js'].loaded) {
        sodBaseAddress = _v_dictSod['sp.js'].url.replace(/sp\.js(\?.+)?$/, '');
    }
    else {
        var scripts = document.getElementsByTagName('script');
        for (var s = 0; s < scripts.length; s++)
            if (scripts[s].src && scripts[s].src.match(/\/sp\.js(\?.+)?$/)) {
                sodBaseAddress = scripts[s].src.replace(/sp\.js(\?.+)?$/, '');
                return sodBaseAddress;
            }
        sodBaseAddress = "/_layouts/15/";
    }
    return sodBaseAddress;
};
var sodDeps = {};
export function registerSodDependency(sod, dep) {
    if (typeof dep === "string") {
        if (_v_dictSod[sod]) {
            SP.SOD.registerSodDep(sod, dep);
            return;
        }
        if (!sodDeps[sod])
            sodDeps[sod] = [];
        sodDeps[sod].push(dep);
    }
    else {
        dep.map(d => registerSodDependency(sod, d));
    }
}
function importer(sod) {
    if (typeof sod === "string") {
        var s = sod.toLowerCase();
        return new Promise((resolve, reject) => {
            if (!_v_dictSod[s] && s != 'sp.ribbon.js') {
                SP.SOD.registerSod(s, getSodBaseAddress() + s);
                for (var d = 0; sodDeps[s] && d < sodDeps[s].length; d++)
                    SP.SOD.registerSodDep(s, sodDeps[s][d]);
            }
            SP.SOD.executeOrDelayUntilScriptLoaded(() => { resolve(); }, s);
            SP.SOD.executeFunc(s, null, null);
        });
    }
    else {
        return Promise.all(sod.map(s => importSod(s)));
    }
}
var spjs_loaded = false;
export function importSod(sod = 'sp.js') {
    if (sod == 'sp.js' && !spjs_loaded) {
        return importer(sod)
            .then(() => {
            spjs_loaded = true;
            registerExtensions();
        });
    }
    return importer(sod);
}
function merge(obj, extension) {
    Object.getOwnPropertyNames(extension.prototype).forEach(name => {
        if (name != "constructor")
            obj.prototype[name] = extension.prototype[name];
    });
}
function registerExtensions() {
    merge(SP.ClientObjectCollection, ClientObjectCollection);
    merge(SP.ClientContext, ClientContext);
    merge(SP.List, List);
    merge(SP["UserCustomActionCollection"], UserCustomActionCollection);
    merge(SP.Guid, Guid);
    SP.SOD['import'] = importSod;
}
export function register(log) {
    return importSod('sp.js')
        .then(() => {
        if (log)
            console.log('SharePoint extensions loaded.');
    });
}
