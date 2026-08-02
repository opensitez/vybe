// vybe-test: js/map_set_deep/set_from_generator
// origin: languages/js/tests/js/test_map_set_deep.rs

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

function* range(n) { for (let i = 0; i < n; i++) yield i; }
const s = new Set(range(5));
console.log(s.size);
console.log([...s].join(","));
