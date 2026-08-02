// vybe-test: js/class_private_advanced/method_chaining_with_private_fields
// origin: languages/js/tests/js/test_class_private_advanced.rs

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

class Query {
    #table = "";
    #conditions = [];
    #limit = null;
    from(table) { this.#table = table; return this; }
    where(cond) { this.#conditions.push(cond); return this; }
    limit(n) { this.#limit = n; return this; }
    build() {
        let q = "SELECT * FROM " + this.#table;
        if (this.#conditions.length > 0) {
            q += " WHERE " + this.#conditions.join(" AND ");
        }
        if (this.#limit !== null) {
            q += " LIMIT " + this.#limit;
        }
        return q;
    }
}
const result = new Query()
    .from("users")
    .where("age > 18")
    .where("active = true")
    .limit(10)
    .build();
__check(__line(result), "SELECT * FROM users WHERE age > 18 AND active = true LIMIT 10");
