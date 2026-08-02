// vybe-test: js/closure_scope_deep/memoize_with_cache_clear
// origin: languages/js/tests/js/test_closure_scope_deep.rs

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

function memoize(fn) {
    const cache = new Map();
    const memo = function(...args) {
        const key = JSON.stringify(args);
        if (!cache.has(key)) cache.set(key, fn(...args));
        return cache.get(key);
    };
    memo.clear = () => cache.clear();
    memo.size = () => cache.size;
    return memo;
}

let calls = 0;
const expensive = memoize((n) => { calls++; return n * n; });
expensive(5); expensive(5); expensive(6);
__check(__line(calls), "2");          // 2 unique calls
__check(__line(expensive.size()), "2"); // 2 cached
expensive.clear();
__check(__line(expensive.size()), "0"); // 0
