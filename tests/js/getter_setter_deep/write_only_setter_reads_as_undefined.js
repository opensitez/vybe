// vybe-test: js/getter_setter_deep/write_only_setter_reads_as_undefined
// origin: languages/js/tests/js/test_getter_setter_deep.rs

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

const obj = {
    _log: [],
    set entry(v) { this._log.push(v); }
};
obj.entry = "a";
obj.entry = "b";
__check(__line(obj.entry), "undefined"); // undefined — no getter
__check(__line(obj._log.join(",")), "a,b");
