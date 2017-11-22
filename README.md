# sharepoint-utilities
A set of convenience methods for SharePoint.

Installing
---

`npm install sharepoint-utilities`.

Extensions
----

This package provides a set of extensions to SharePoint's JS Object Model to make development slightly less tedious.

```typescript
// A set of convenient extension methods
declare namespace SP {
    export interface SOD {
        import(sod: string | string[]): Promise<any>;
    }

    export interface ClientObjectCollection<T> extends IEnumerable<T> {
    }

    export interface ClientRuntimeContext {
        /** A shorthand for context.executeQueryAsync except wrapped as a JS Promise object */
        executeQuery(): Promise<ClientRequestSucceededEventArgs>;
    }

    export interface List {
        /** A shorthand to list.getItems with just the queryText and doesn't require a SP.CamlQuery to be constructed
        @param queryText the queryText to use for the query.set_ViewXml() call */
        get_queryResult(queryText: string): ListItemCollection;
    }

    namespace Guid {
        export function generateGuid(): string;
    }
}

// Collection methods similar to those available from lodash
declare interface IEnumerable<T> {
    getEnumerator(): IEnumerator<T>;

    /** Execute a callback for every element in the matched set.
    @param callback The function that will called for each element, and passed an index and the element itself */
    each?(callback: (index?: number, item?: T) => void): void;

    /**
     * Creates an array of values by running each element in collection through iteratee.
     *
     * @param iteratee The function invoked per iteration.
     * @return Returns the new mapped array. */
    map?<TResult>(iteratee: (item?: T, index?: number) => TResult): TResult[];

    /** Converts a collection to a regular JS array. */
    toArray?(): T[];

    /** Callback for some collection method
     * @callback iterateeCallback
     * @param item The current element being processed in the collection
     * @param {number} index The index of the current element being processed in the collection
     * @param collection The collection some() was called upon.
     */

    /** Tests whether at least one element in the collection passes the test implemented by the provided function.
     * @param {iterateeCallback} iteratee Function to test for each element in the collection
     * @returns true if the callback */
    some?(iteratee?: (item: T, index?: number, collection?: IEnumerable<T>) => boolean): boolean;

    /** Tests whether at least one element in the collection passes the test implemented by the provided function.
     * @param {iterateeCallback} iteratee Function to test for each element in the collection
     * @returns true if the callback */
    every?(iteratee?: (item: T, index?: number, collection?: IEnumerable<T>) => boolean): boolean;

    /** Tests whether at least one element in the collection passes the test implemented by the provided function.
     * @param {iterateeCallback} iteratee Function to execute on each element in the collection
     * @returns true if the callback */
    find?(iteratee?: (item: T, index?: number, collection?: IEnumerable<T>) => boolean): T;

    /** Returns the first element in the collection or null if none
     * @param iteratee An optional function to filter by
     * @return Returns the first item in the collection */
    firstOrDefault?(iteratee?: (item?: T) => boolean): T;
}
```

Example Usage:
---

```javascript
/* instead of:
SP.SOD.executeOrDelayUntilScriptLoaded(function() {
    ...
}, 'sp.js');
LoadSodByKey('sp.js');
*/
SP.SOD.import('sp.js')
.then(function() {
    var context = new SP.ClientContext();
    var web = context.get_web();
    var lists = web.get_lists();
    var list = lists.getByTitle('list_title');

    /* instead of:
    var query = new SP.CamlQuery();
    query.set_viewXml('<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">Hello world!</Value></Eq></Where></Query></View>');
    var items = list.getItems(query);
    */
    var items = list.get_queryResult('<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">Hello world!</Value></Eq></Where></Query></View>');

    context.load(items);

    /* instead of:
    context.executeQueryAsync(
        Function.createDelegate(this, function(sender, args) {
            ...
        }),
        Function.createDelegate(this, function(sender, args) {
            ...
        })
    );
    */
    context.executeQuery()
    .then(function(args) {
        /* instead of:
        var transformed = [];
        var enumerator = items.getEnumerator();
        while(enumerator.moveNext()) {
            var item = enumerator.get_current();
            transformed.push({ title: item.get_title() });
        }
        */
        var transformed = items.map(function(item, i) {
            return { title: item.get_title() };
        });

        console.log(transformed);
    })
    .catch(function(args) {
        console.log(args.get_message());
    });
});
```
