// vybe-test: js/tagged_template_literal_raw_strings/test_js_tagged_template_sql_parameterized_query_builder
// origin: languages/js/tests/js/test_js_tagged_template_literal_raw_strings.rs

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

function sql(strings, ...values) {
    const query = strings.join("?");
    return query + "|Params=" + values.join(",");
}
const id = 42, status = "active";
__check(__line(sql`SELECT * FROM users WHERE id = ${id} AND status = ${status}`), "SELECT * FROM users WHERE id = ? AND status = ?|Params=42,active");
