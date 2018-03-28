# sharepoint-utilities
A set of convenience methods for SharePoint.

Installing
---
`npm install sharepoint-utilities`

Usage
-----

```typescript
import {register} from 'sharepoint-utilities';
register().then(() => {
});
```

Browser Console Usage
----

```typescript
// Only works in browser with support for ES6 modules:
import("https://dennispg.github.io/sharepoint-utilities/es6/index.js")
.then(m=>m.register(true));
```

Extensions
----
This package provides a set of extensions to SharePoint's JavaScript Object Model to make development slightly less tedious. These include collection methods similar to those available from `lodash` (https://lodash.com/docs).

Reference
----

### Table of contents
-   [`SP.SOD.import`](#spsodimport)
-   [`SP.ClientContext.executeQuery`](#spclientcontextexecutequery)
-   [`SP.List.get_queryResult`](#splistget_queryresult)
-   [`SP.Guid.generateGuid`](#spguidgenerateguid)
-   [`IEnumerable<T>.each`](#ienumerableteach)
-   [`IEnumerable<T>.every`](#ienumerabletevery)
-   [`IEnumerable<T>.find`](#ienumerabletfind)
-   [`IEnumerable<T>.filter`](#ienumerabletfilter)
-   [`IEnumerable<T>.firstOrDefault`](#ienumerabletfirstordefault)
-   [`IEnumerable<T>.forEach`](#ienumerabletforeach)
-   [`IEnumerable<T>.groupBy`](#ienumerabletgroupby)
-   [`IEnumerable<T>.map`](#ienumerabletmap)
-   [`IEnumerable<T>.reduce`](#ienumerabletreduce)
-   [`IEnumerable<T>.some`](#ienumerabletsome)
-   [`IEnumerable<T>.toArray`](#ienumerablettoarray)
-   [`SP.UserCustomActionCollection.ensure`](#spusercustomactioncollectionensure)
-   [`SP.UserCustomActionCollection.remove`](#spusercustomactioncollectionremove)
-   [`SP.UserCustomActionCollection.removeByTitle`](#spusercustomactioncollectionremoveByTitle)
-   [`SP.UserCustomActionCollection.removeByName`](#spusercustomactioncollectionremoveByName)
-   [`SP.UserCustomActionCollection.removeById`](#spusercustomactioncollectionremoveById)

## SP.SOD.import
A wrapper around SharePoint's Script-On-Demand using Promises.

**Parameters**
-   `sod` **`(string | string[])`** Script or scripts to load from SharePoint

**Example**
```javascript
// instead of:
SP.SOD.executeOrDelayUntilScriptLoaded(function() {
    <...>
}, 'sp.js');
LoadSodByKey('sp.js');
```
```javascript
// use this:
SP.SOD.import('sp.js')
.then(function() {
    <...>
});
```

## SP.ClientContext.executeQuery
A shorthand for context.executeQueryAsync except wrapped as a Promise object

**Example**
```javascript
// instead of:
var context = new SP.ClientContext();
context.executeQueryAsync(
    Function.createDelegate(this, function(sender, args) {
        <...>
    }),
    Function.createDelegate(this, function(sender, args) {
        <...>
    })
);
```
```javascript
// use this:
var context = new SP.ClientContext();
context.executeQuery()
.then(function(args) {
    <...>
});
```

## SP.List.get_queryResult
Returns a collection of items from the list based on the specified query. A shorthand for SP.List.getItems with just the queryText and doesn't require a SP.CamlQuery to be constructed

**Parameters**
-   `queryText` **(string)** The CAML query to call SP.List.getItems with.

Returns an **SP.ListItemCollection**

**Example**
```javascript
// instead of:
var query = new SP.CamlQuery();
query.set_viewXml('<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">Hello world!</Value></Eq></Where></Query></View>');
var items = list.getItems(query);
```
```typescript
/// use this:
var items = list.get_queryResult('<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">Hello world!</Value></Eq></Where></Query></View>');
```

## SP.Guid.generateGuid

**Example**
```typescript
let guid = SP.Guid.generateGuid();
```

## IEnumerable<T>.each
Execute a callback for every element in the matched set. (use jQuery style callback signature)

**Example**
```typescript
// instead of:
let items = list.getItems();
var enumerator = items.getEnumerator();
while(enumerator.moveNext()) {
    var item = enumerator.get_current();
    console.log(item.get_title());
}
```
```typescript
// use this:
let items = list.getItems();
items.each(function(i, item) {
    console.log(item.get_title());
});
```

## IEnumerable<T>.every
Tests whether every element in the collection passes the test implemented by the provided function.

**Example**
```typescript
let items = list.getItems();
items.every(item => item.get_item('Title') == "") == false;
```

## IEnumerable<T>.find
Finds the first element in the collection which passes the test implemented by the provided function, or null if not found.

**Example**
```typescript
let items = list.getItems();
var found_item = items.find(item => item.get_item('Title') == "");
```

## IEnumerable<T>.filter
Iterates over elements of collection, returning an array of all elements predicate returns truthy for. The predicate is invoked with three arguments: *(value, index|key, collection)*.

**See**

[lodash filter function documentation](https://lodash.com/docs/#filter)

**Example**
```typescript
let items = list.getItems();

items.filter(item => item.get_item('Title') == "")
// => all list items with empty Title field

items.filter({ "Title": "", "Author": web.get_currentUser() });
// => all list items where Title is empty and the item was created by the current user

items.filter('IsActive');
// => all items where IsActive field is truthy (non-null, not false, not 0)
```

## IEnumerable<T>.firstOrDefault
Returns the first element in the collection or null if none

**Example**
```typescript
let items = list.getItems();
var item = items.firstOrDefault();
if(item != null)
    console.log('Got the first item in the collection.')
else
    console.log('The collection was empty.')
```

## IEnumerable<T>.forEach
Execute a callback for every element in the matched set. (uses array.forEach style callback signature)

**Example**
```typescript
// instead of:
let items = list.getItems();
var enumerator = items.getEnumerator();
while(enumerator.moveNext()) {
    var item = enumerator.get_current();
    console.log(item.get_title());
}
```
```typescript
// use this:
let items = list.getItems();
items.each(function(item, i) {
    console.log(item.get_title());
});
```

## IEnumerable<T>.groupBy
Creates an object composed of keys generated from the results of running each element of `collection` thru `iteratee`. The order of grouped values is determined by the order they occur in `collection`. The corresponding value of each key is an array of elements responsible for generating the key. The iteratee is invoked with three arguments: (value, index, collection).

**Example**
```typescript
let items = list.getItems();
var groups = items.groupBy(function(item) {
    return item.get_item('ContentType');
});
// => { 'Document': [item1, item2], 'Page': [item3, item4] }
```

## IEnumerable<T>.map
Creates an array of values by running each element in collection through iteratee.

**Example**
```typescript
// instead of:
let items = list.getItems();
var transformed = [];
var enumerator = items.getEnumerator();
while(enumerator.moveNext()) {
    var item = enumerator.get_current();
    transformed.push({ title: item.get_title() });
}
```
```typescript
// use this:
let items = list.getItems();
var transformed = items.map(function(item, i) {
    return { title: item.get_title() };
});
```

## IEnumerable<T>.reduce
Reduces `collection` to a value which is the accumulated result of running each element in `collection` thru `iteratee`, where each successive invocation is supplied the return value of the previous. If `accumulator` is not given, the first element of `collection` is used as the initial value. The iteratee is invoked with four arguments: (accumulator, value, index|key, collection).

**Example**
```typescript
let items = list.getItems();
var total_sum = items.reduce(function(sum, item) {
    return sum + item.get_item('Total');
}, 0)
```

## IEnumerable<T>.some
Tests whether at least one element in the collection passes the test implemented by the provided function.

**Example**
```typescript
let items = list.getItems();
if (items.some((item, i, items) => item.get_item('Title') == ""))
    console.log('Found an empty title.');
```

## IEnumerable<T>.toArray
Converts a collection to regular **Array**

**Example**
```typescript
let items = list.getItems();
let item_array = items.toArray();
```

## SP.UserCustomActionCollection.ensure
Adds a `SP.UserCustomAction` object to the collection. If an object with the same name or title already, it is overwritten instead of creating a duplicate.

**IMPORTANT NOTE**: This operation triggers a SP.ClientContext.executeQuery operation.

**Example**
```typescript
var web = context.get_web();
var actions = web.get_userCustomActions();
actions.ensure({
    title: "ExampleCustomAction",
    scriptSrc: "~sitecollection/Style Library/custom_script.js",
    location: "ScriptLink",
    sequence: 1
})
.then(id => {
    console.log(`ScriptLink custom action with ID:${id} as added successfully.`);
});
```

## SP.UserCustomActionCollection.remove
Deletes all `SP.UserCustomAction` objects from the collection where the provided selector function returns true.

## SP.UserCustomActionCollection.removeByTitle
Deletes all `SP.UserCustomAction` object from the collection with matching title value.

**Example**
```typescript
var web = context.get_web();
var actions = web.get_userCustomActions();
actions.removeByTitle("ExampleCustomAction");
```

## SP.UserCustomActionCollection.removeByName
Deletes all `SP.UserCustomAction` object from the collection with matching name value.

## SP.UserCustomActionCollection.removeById
Deletes all `SP.UserCustomAction` object from the collection with matching ID value.