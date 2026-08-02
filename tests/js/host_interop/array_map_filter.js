// vybe-test: js/host_interop/array_map_filter
// origin: languages/js/tests/js/js_host_interop_test.rs

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

let a = [1, 2, 3, 4, 5];
        let evens = a.filter(x => x % 2 === 0);
        let doubled = evens.map(x => x * 2);
        __check(__line(doubled.join(",")), "4,8");
