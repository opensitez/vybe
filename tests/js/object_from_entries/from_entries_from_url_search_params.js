// vybe-test: js/object_from_entries/from_entries_from_url_search_params
// origin: languages/js/tests/js/test_object_from_entries.rs

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

function parseParams(qs) {
    const obj = {};
    for (const pair of qs.split("&")) {
        const [k, v] = pair.split("=");
        obj[k] = v;
    }
    return obj;
}
const obj = parseParams("a=1&b=2&c=3");
console.log(obj.a);
console.log(obj.b);
console.log(obj.c);
