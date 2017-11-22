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
declare global {
    namespace SP {
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
}
export function registerSodDependency(sod: string, dep: string | string[]);
export function importSod(sod: string | string[]): Promise<any>;
export function registerExtensions();