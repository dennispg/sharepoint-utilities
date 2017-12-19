type iterateeFunction<T> = (item?: T, index?: number, collection?: SP.ClientObjectCollection<T>) => boolean | void;
type filterPredicate<T> = iterateeFunction<T> | {[prop: string]:any} | string | string[];

declare interface IEnumerable<T> {
    getEnumerator(): IEnumerator<T>;

    /** Execute a callback for every element in the matched set. (with a jQuery style callback signature)
    @param callback The function that will called for each element, and passed an index and the element itself
    @deprecated use forEach instead! */
    each?(callback: (index?: number, item?: T) => boolean | void): void;

    /** Execute a callback for every element in the matched set.
    @param iteratee The function that will called for each element, and passed an index and the element itself */
    forEach?(iteratee: iterateeFunction<T>): void;

    /**
     * Creates an array of values by running each element in collection through iteratee.
     *
     * @param iteratee The function invoked per iteration.
     * @return Returns the new mapped array. */
    map?<TResult>(iteratee: (item?: T, index?: number, coll?: SP.ClientObjectCollection<T>) => TResult): TResult[];

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
    some?(iteratee?: iterateeFunction<T>): boolean;

    /** Tests whether at least one element in the collection passes the test implemented by the provided function.
     * @param {iterateeCallback} iteratee Function to test for each element in the collection
     * @returns true if the callback */
    every?(iteratee?: iterateeFunction<T>): boolean;

    /** Tests whether at least one element in the collection passes the test implemented by the provided function.
     * @param {iterateeCallback} iteratee Function to execute on each element in the collection
     * @returns true if the callback */
    find?(iteratee?: iterateeFunction<T>): T;

    /**
     * Creates a function that performs a partial deep comparison between a given
     * object and `source`, returning `true` if the given object has equivalent
     * property values, else `false`.
     *
     * @param {Object} source The object of property values to match.
     * @returns {Function} Returns the new spec function.
     */
    matches?(source: {[prop: string]:any}): iterateeFunction<T>

    /**
     * Iterates over elements of `array`, returning an array of all elements
     * `predicate` returns truthy for. The predicate is invoked with three
     * arguments: (value, index, array).
     *
     * @param {Function} predicate The function invoked per iteration.
     * @returns {Array} Returns the new filtered array.
     */
    filter?(predicate: filterPredicate<T>): T[];

    /** Returns the first element in the collection or null if none
     * @param iteratee An optional function to filter by
     * @return Returns the first item in the collection */
    firstOrDefault?(iteratee?: iterateeFunction<T>): T;

    /**
     * Reduces `collection` to a value which is the accumulated result of running
     * each element in `collection` thru `iteratee`, where each successive
     * invocation is supplied the return value of the previous.
     * The iteratee is invoked with four arguments: (accumulator, value, index|key, collection)
     *
     * @param {Function} iteratee The function invoked per iteration.
     * @param {*} [accumulator] The initial value.
     * @returns {*} Returns the accumulated value.
     */
    reduce?<TResult>(
        iteratee: (prev: TResult, curr: T, index: number, list: SP.ClientObjectCollection<T>) => TResult,
        accumulator: TResult
    ): TResult;

    /**
     * Creates an object composed of keys generated from the results of running
     * each element of `collection` thru `iteratee`. The order of grouped values
     * is determined by the order they occur in `collection`. The corresponding
     * value of each key is an array of elements responsible for generating the
     * key. The iteratee is invoked with three arguments: (value, index, collection).
     *
     * @param {Function} iteratee The iteratee to transform keys.
     * @returns {Object} Returns the composed aggregate object.
     */
    groupBy?(iteratee?: (value: T, index?: number, collection?: SP.ClientObjectCollection<T>) => string | number): {[group:string]:T[]};
}

declare global {
    namespace SP {
        export interface ClientObjectCollection<T> extends IEnumerable<T> {
        }

        export interface ClientRuntimeContext {
            /** A shorthand for context.executeQueryAsync except wrapped as a JS Promise object */
            executeQuery(): Promise<ClientRequestSucceededEventArgs>;
        }

        export interface List {
            /** A shorthand to list.getItems with just the queryText and doesn't require a SP.CamlQuery to be constructed
            @param queryText the queryText to use for the query.set_ViewXml() call */
            get_queryResult<T=any>(queryText: string): ListItemCollection<T>;
        }

        namespace Guid {
            export function generateGuid(): string;
        }

        export interface SOD {
            import(sod: string | string[]): Promise<any>;
        }
    }
}
export function registerSodDependency(sod: string, dep: string | string[]);
export function importSod(sod: string | string[]): Promise<any>;
export function register();