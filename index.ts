type iterateeFunction<T> = (item?: T, index?: number, collection?: SP.ClientObjectCollection<T>) => boolean | void;
type filterPredicate<T> = iterateeFunction<T> | {[prop: string]:any} | string | string[];

class ClientObjectCollection<T> implements IEnumerable<T> {
    getEnumerator: () => IEnumerator<T>;

    /** Execute a callback for every element in the matched set. (with a jQuery style callback signature)
    @param callback The function that will called for each element, and passed an index and the element itself
    @deprecated use forEach instead! */
    each?(callback: (index?: number, item?: T) => boolean | void): void {
        return this.forEach((item, i) => {
            callback(i, item);
        });
    }

    /** Execute a callback for every element in the matched set.
    @param iteratee The function that will called for each element, and passed an index and the element itself */
    forEach?(iteratee: iterateeFunction<T>): void {
        var index = 0, enumerator = this.getEnumerator();
        while (enumerator.moveNext()) {
            if (iteratee(enumerator.get_current(), index++, <SP.ClientObjectCollection<T>><any>this) === false)
                break;
        }
    }

    /**
     * Creates an array of values by running each element in collection through iteratee.
     *
     * @param iteratee The function invoked per iteration.
     * @return Returns the new mapped array. */
    map?<TResult>(iteratee: (item?: T, index?: number, coll?: SP.ClientObjectCollection<T>) => TResult): TResult[] {
        var index = -1, enumerator = this.getEnumerator(), result = [];
        while (enumerator.moveNext()) {
            result[++index] = iteratee(enumerator.get_current(), index, <SP.ClientObjectCollection<T>><any>this);
        }
        return result;
    }

    /** Converts a collection to a regular JS array. */
    toArray?(): T[] {
        var collection: T[] = [];
        this.each((i, item: T) => {
            collection.push(item);
        });
        return collection;
    }

    /** Callback for some collection method
     * @callback iterateeCallback
     * @param item The current element being processed in the collection
     * @param {number} index The index of the current element being processed in the collection
     * @param collection The collection some() was called upon.
     */

    /** Tests whether at least one element in the collection passes the test implemented by the provided function.
     * @param {iterateeCallback} iteratee Function to test for each element in the collection
     * @returns true if the callback */
    some?(iteratee?: iterateeFunction<T>): boolean {
        var val = false;
        this.each((i, item) => {
            if(iteratee(item, i, <SP.ClientObjectCollection<T>><any>this)) {
                val = true;
                return false;
            }
        });
        return val;
    }

    /** Tests whether at least one element in the collection passes the test implemented by the provided function.
     * @param {iterateeCallback} iteratee Function to test for each element in the collection
     * @returns true if the callback */
    every?(iteratee?: iterateeFunction<T>): boolean {
        var val = true;
        var hasitems = false;
        this.each((i, item) => {
            hasitems = true;
            if(!iteratee(item, i, <SP.ClientObjectCollection<T>><any>this)) {
                val = false;
                return false;
            }
        });
        return hasitems && val;
    }

    /** Tests whether at least one element in the collection passes the test implemented by the provided function.
     * @param {iterateeCallback} iteratee Function to execute on each element in the collection
     * @returns true if the callback */
    find?(iteratee?: iterateeFunction<T>): T {
        var val = undefined;
        this.each((i, item) => {
            if(iteratee(item, i, <SP.ClientObjectCollection<T>><any>this)) {
                val = item;
                return false;
            }
        });
        return val;
    }

    reduce?<TResult>(
        iteratee: (prev: TResult, curr: T, index: number, list: SP.ClientObjectCollection<T>) => TResult,
        accumulator: TResult
    ): TResult {
        this.forEach((item, i) => {
            accumulator = iteratee(accumulator, item, i, <SP.ClientObjectCollection<T>><any>this);
        });
        return accumulator;
    }

    groupBy?(iteratee?: (value: T, index?: number, collection?: SP.ClientObjectCollection<T>) => string | number): {[group:string]:T[]} {
        return this.reduce((result, value) => {
            var key = iteratee(value)
            if (Object.prototype.hasOwnProperty.call(result, key))
                result[key].push(value)
            else
                result[key] = [value]
            return result
        }, {})
    }

    matches?<T>(source: {[prop: string]:any}): iterateeFunction<T> {
        return item => {
            for(var prop in (<{[prop: string]:any}>source)) {
                var compare_val = source[prop];
                var value = null;
                if(item instanceof SP.ListItem) {
                    value = item.get_item(prop);
                    if(value instanceof SP.FieldLookupValue || value instanceof SP.FieldUserValue) {
                        value = value.get_lookupId();
                        if(compare_val instanceof SP.User)
                            compare_val = compare_val.get_id();
                        else if(compare_val instanceof SP.FieldLookupValue || compare_val instanceof SP.FieldUserValue)
                            compare_val = compare_val.get_lookupId();
                    }
                    else if(value instanceof SP.FieldUrlValue) {
                        value = value.get_url() + ';#' + value.get_description();
                        if(compare_val instanceof SP.FieldUrlValue)
                            compare_val = compare_val.get_url() + ';#' + compare_val.get_description();
                    }
                }
                else value = item[prop]

                return value == compare_val;
            }
        }
    }

    property?<T>(props: string | string[]): iterateeFunction<T> {
        var properties: string[] = Array.isArray(props) ? props : [props];

        return item => {
            if(item instanceof SP.ListItem)
                return properties.every(prop => item.get_item(prop));
            else
                return properties.every(prop => item[<string> prop]);
        }
    }

    filter?(predicate: filterPredicate<T>): T[] {
        var predicateType = Object.prototype.toString.call(predicate);
        switch(predicateType) {
            case '[object Function]':
                var items = [];
                var filter = <iterateeFunction<T>>predicate;
                this.forEach((item, i) => {
                    if(filter(item, i, <SP.ClientObjectCollection<T>><any>this))
                        items.push(item);
                });
                return items;

            case 'object':
                return this.filter(this.matches(<{[prop: string]:any}> predicate));

            case 'string':
            case '[object Array]':
                return this.filter(this.property(<string|string[]> predicate));
        }
        return null;
    }

    /** Returns the first element in the collection or null if none
     * @param iteratee An optional function to filter by
     * @return Returns the first item in the collection */
    firstOrDefault?(iteratee?: iterateeFunction<T>): T {
        var index = 0, enumerator = this.getEnumerator();
        if (enumerator.moveNext()) {
            var current = enumerator.get_current();
            if(iteratee) {
                if(iteratee(current, index++, <SP.ClientObjectCollection<T>><any>this))
                    return current;
            }
            else
                return current;
        }
        return null;
    }
}

class ClientContext {
    /** A shorthand for context.executeQueryAsync except wrapped as a JS Promise object */
    executeQuery() {
        var context = <SP.ClientRuntimeContext><any>this;
        return new Promise<SP.ClientRequestSucceededEventArgs>((resolve, reject) => {
            context.executeQueryAsync(
                (sender, args: SP.ClientRequestSucceededEventArgs) => { resolve(args); },
                (sender, args: SP.ClientRequestFailedEventArgs) => { reject(args); }
            );
        });
    }
}

class List {
    /** A shorthand to list.getItems with just the queryText and doesn't require a SP.CamlQuery to be constructed
    @param queryText the queryText to use for the query.set_ViewXml() call */
    get_queryResult(queryText: string): SP.ListItemCollection {
        var query = new SP.CamlQuery();
        query.set_viewXml(queryText);
        return (<SP.List><any>this).getItems(query);
    }
}

class Guid {
    static generateGuid() {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
            var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    };
}

declare var _v_dictSod : { [address: string]: any };

var sodBaseAddress = null;
var getSodBaseAddress = function() {
    if(sodBaseAddress)
        return sodBaseAddress;

    if(_v_dictSod['sp.js'] && _v_dictSod['sp.js'].loaded) {
        sodBaseAddress = _v_dictSod['sp.js'].url.replace(/sp\.js(\?.+)?$/,'');
    }
    else {
        var scripts = document.getElementsByTagName('script');
        for(var s=0; s < scripts.length; s++)
            if(scripts[s].src && scripts[s].src.match(/\/sp\.js(\?.+)?$/)) {
                sodBaseAddress = scripts[s].src.replace(/sp\.js(\?.+)?$/,'');
                return sodBaseAddress;
            }
        sodBaseAddress = "/_layouts/15/";
    }

    return sodBaseAddress;
}

var sodDeps = {};
export function registerSodDependency(sod: string, dep: string | string[]) {
    if(typeof dep === "string") {
        if(_v_dictSod[sod]) {
            SP.SOD.registerSodDep(sod, dep);
            return;
        }
        if(!sodDeps[sod])
            sodDeps[sod] = [];
        sodDeps[sod].push(dep);
    }
    else {
        dep.map(d => registerSodDependency(sod, d));
    }
}

function importer(sod: string | string[]): Promise<any> {
    if(typeof sod === "string") {
        var s = sod.toLowerCase();
        return new Promise<any>((resolve, reject) => {
            if (!_v_dictSod[s] && s != 'sp.ribbon.js') {
                // if its not registered, we can only assume it needs to be
                SP.SOD.registerSod(s, getSodBaseAddress() + s);
                for(var d=0;sodDeps[s] && d<sodDeps[s].length;d++)
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
export function importSod(sod: string | string[] = 'sp.js'): Promise<any> {
    if(sod == 'sp.js' && !spjs_loaded) {
        return importer(sod)
        .then(() => {
            spjs_loaded = true;
            registerExtensions();
        })
    }
    return importer(sod);
}

function merge(obj, extension) {
    Object.getOwnPropertyNames(extension.prototype).forEach(name => {
        obj.prototype[name] = extension.prototype[name];
    });
}

function registerExtensions() {
    merge(SP.ClientObjectCollection, ClientObjectCollection);
    merge(SP.ClientContext, ClientContext);
    merge(SP.List, List);
    merge(SP.Guid, Guid);
    SP.SOD['import'] = importSod;
}

export function register() {
    return importSod('sp.js');
}