# sharepoint-utilities
A set of convenience methods for SharePoint.

Installing
---
`npm install sharepoint-utilities`.

Extensions
----
This package provides a set of extensions to SharePoint's JavaScript Object Model to make development slightly less tedious. These include collection methods similar to those available from `lodash`.

Reference
----

### Table of contents
-   [`SP.SOD.import`](#spsodimport)
-   [`SP.ClientContext.executeQuery`](#spclientcontextexecutequery)
-   [`SP.List.get_queryResult`](#splistget_queryresult)
-   [`SP.Guid.generateGuid`](#spguidgenerateguid)
-   [`SP.ClientObjectCollection<T>.map`](#spclientobjectcollectioneach)
-   [`SP.ClientObjectCollection<T>.every`](#spclientobjectcollectionevery)
-   [`SP.ClientObjectCollection<T>.find`](#spclientobjectcollectionfind)
-   [`SP.ClientObjectCollection<T>.firstOrDefault`](#spclientobjectcollectionfirstordefault)
-   [`SP.ClientObjectCollection<T>.map`](#spclientobjectcollectionmap)
-   [`SP.ClientObjectCollection<T>.some`](#spclientobjectcollectionsome)
-   [`SP.ClientObjectCollection<T>.toArray`](#spclientobjectcollectiontoarray)

## SP.SOD.import
A wrapper around SharePoint's Script-On-Demand using Promises.

**Parameters**
-   `sod` **(string | string[])** Script or scripts to load from SharePoint

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

## SP.ClientObjectCollection<T>.each
Execute a callback for every element in the matched set.

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
## SP.ClientObjectCollection<T>.every
Tests whether every element in the collection passes the test implemented by the provided function.

**Example**
```typescript
let items = list.getItems();
items.every(item => item.get_item('Title') == "") == false;
```

## SP.ClientObjectCollection<T>.find
Finds the first element in the collection which passes the test implemented by the provided function, or null if not found.

**Example**
```typescript
let items = list.getItems();
var found_item = items.find(item => item.get_item('Title') == "");
```

## SP.ClientObjectCollection<T>.firstOrDefault
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

## SP.ClientObjectCollection<T>.map
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

## SP.ClientObjectCollection<T>.some
Tests whether at least one element in the collection passes the test implemented by the provided function.

**Example**
```typescript
let items = list.getItems();
if (items.some((item, i, items) => item.get_item('Title') == ""))
    console.log('Found an empty title.');
```

## SP.ClientObjectCollection<T>.toArray
Converts a collection to regular **Array**

**Example**
```typescript
let items = list.getItems();
let item_array = items.toArray();
```