// vybe-test: js/object_immutability/deep_freeze_pattern
// origin: languages/js/tests/js/test_object_immutability.rs

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

function deepFreeze(obj) {
    Object.getOwnPropertyNames(obj).forEach(key => {
        const val = obj[key];
        if (typeof val === "object" && val !== null) deepFreeze(val);
    });
    return Object.freeze(obj);
}
const config = deepFreeze({ db: { host: "localhost", port: 5432 } });
config.db.port = 9999;
console.log(config.db.port);
