// vybe-test: js/chaining_patterns/sql_builder_chaining
// origin: languages/js/tests/js/test_chaining_patterns.rs

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

class SQL {
    #parts = [];
    select(...cols) { this.#parts.push(`SELECT ${cols.join(", ")}`); return this; }
    from(table) { this.#parts.push(`FROM ${table}`); return this; }
    where(cond) { this.#parts.push(`WHERE ${cond}`); return this; }
    orderBy(col) { this.#parts.push(`ORDER BY ${col}`); return this; }
    limit(n) { this.#parts.push(`LIMIT ${n}`); return this; }
    build() { return this.#parts.join(" "); }
}
const query = new SQL()
    .select("id", "name")
    .from("users")
    .where("active = 1")
    .orderBy("name")
    .limit(10)
    .build();
__check(__line(query), "SELECT id, name FROM users WHERE active = 1 ORDER BY name LIMIT 10");
