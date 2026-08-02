// vybe-test: js/immutable_patterns/record_update_pattern_with_class
// origin: languages/js/tests/js/test_immutable_patterns.rs

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

class Record {
    constructor(data) { Object.assign(this, data); Object.freeze(this); }
    update(changes) { return new Record({ ...this, ...changes }); }
}
const r1 = new Record({ x: 1, y: 2, z: 3 });
const r2 = r1.update({ y: 99 });
__check(__line(r1.y), "2");
__check(__line(r2.y), "99");
__check(__line(r2.x), "1");
