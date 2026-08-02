// vybe-test: js/tagged_template_deep/tagged_template_sql_builder
// origin: languages/js/tests/js/test_tagged_template_deep.rs

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
    const query = strings.reduce((acc, s, i) => acc + s + (i < values.length ? "?" : ""), "");
    return { query, params: values };
}
const id = 42, name = "Alice";
const result = sql`SELECT * FROM users WHERE id = ${id} AND name = ${name}`;
__check(__line(result.params.length), "2");
__check(__line(result.params[0]), "42");
__check(__line(result.params[1]), "Alice");
