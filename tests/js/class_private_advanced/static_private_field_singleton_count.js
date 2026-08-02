// vybe-test: js/class_private_advanced/static_private_field_singleton_count
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

class Connection {
    static #pool = [];
    static #maxSize = 3;
    id;
    constructor(id) { this.id = id; }
    static acquire(id) {
        if (Connection.#pool.length < Connection.#maxSize) {
            const conn = new Connection(id);
            Connection.#pool.push(conn);
            return conn;
        }
        return null;
    }
    static poolSize() { return Connection.#pool.length; }
}
Connection.acquire("a");
Connection.acquire("b");
Connection.acquire("c");
const d = Connection.acquire("d");
__check(__line(Connection.poolSize()), "3");
__check(__line(d), "null");
