// vybe-test: js/string_processing_deep/parse_query_string
// origin: languages/js/tests/js/test_string_processing_deep.rs

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

function parseQS(qs) {
    return Object.fromEntries(
        qs.split("&").map(p => {
            const [k, v] = p.split("=");
            return [decodeURIComponent(k), decodeURIComponent(v ?? "")];
        })
    );
}
const q = parseQS("name=Alice&age=30&city=New%20York");
__check(__line(q.name), "Alice");
__check(__line(q.age), "30");
__check(__line(q.city), "New York");
