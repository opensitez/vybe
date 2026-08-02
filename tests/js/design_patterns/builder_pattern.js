// vybe-test: js/design_patterns/builder_pattern
// origin: languages/js/tests/js/test_design_patterns.rs

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
    constructor() { this._table = ""; this._where = []; this._limit = null; }
    from(table) { this._table = table; return this; }
    where(cond) { this._where.push(cond); return this; }
    limit(n) { this._limit = n; return this; }
    build() {
        let q = `SELECT * FROM ${this._table}`;
        if (this._where.length) q += ` WHERE ${this._where.join(" AND ")}`;
        if (this._limit) q += ` LIMIT ${this._limit}`;
        return q;
    }
}
const query = new QueryBuilder()
    .from("users")
    .where("age > 18")
    .where("active = 1")
    .limit(10)
    .build();
__check(__line(query), "SELECT * FROM users WHERE age > 18 AND active = 1 LIMIT 10");
