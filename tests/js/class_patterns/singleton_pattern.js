// vybe-test: js/class_patterns/singleton_pattern
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

class Database {
    static instance = null;
    constructor(name) { this.name = name; }
    static getInstance(name) {
        if (!Database.instance) {
            Database.instance = new Database(name);
        }
        return Database.instance;
    }
}
let db1 = Database.getInstance("main");
let db2 = Database.getInstance("other");
__check(__line(db1 === db2), "true");
__check(__line(db2.name), "main");
