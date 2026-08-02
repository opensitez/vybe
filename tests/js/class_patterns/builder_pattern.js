// vybe-test: js/class_patterns/builder_pattern
// origin: languages/js/tests/js/test_class_patterns.rs

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

class QueryBuilder {
    constructor() { this.parts = []; }
    select(fields) { this.parts.push("SELECT " + fields); return this; }
    from(table) { this.parts.push("FROM " + table); return this; }
    where(cond) { this.parts.push("WHERE " + cond); return this; }
    build() { return this.parts.join(" "); }
}
let q = new QueryBuilder()
    .select("*")
    .from("users")
    .where("age > 18")
    .build();
__check(__line(q), "SELECT * FROM users WHERE age > 18");
