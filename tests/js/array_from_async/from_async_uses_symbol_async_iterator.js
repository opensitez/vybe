// vybe-test: js/array_from_async/from_async_uses_symbol_async_iterator
// origin: languages/js/tests/js/test_array_from_async.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

const asyncIterable = {
    [Symbol.asyncIterator]() {
        let i = 0;
        const vals = [100, 200, 300];
        return {
            async next() {
                if (i < vals.length) return { value: vals[i++], done: false };
                return { done: true };
            }
        };
    }
};
Array.fromAsync(asyncIterable).then(arr => console.log(arr.join(",")));
