// vybe-test: js/url_builtin_matrix/url_searchparams_iteration_order_matches_query
// origin: languages/js/tests/js/test_url_builtin_matrix.rs

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

const u = new URL("https://example.com/a?b=1&a=2&b=3");
const out = [];
for (const [k, v] of u.searchParams) out.push(k + "=" + v);
console.log(out.join(","));
