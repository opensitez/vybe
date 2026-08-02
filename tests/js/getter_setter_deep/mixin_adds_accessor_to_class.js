// vybe-test: js/getter_setter_deep/mixin_adds_accessor_to_class
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

const Timestamped = (Base) => class extends Base {
    get timestamp() { return this._ts || 0; }
    set timestamp(v) { this._ts = v; }
};

class Record {}
class TimedRecord extends Timestamped(Record) {}

const r = new TimedRecord();
r.timestamp = 12345;
__check(__line(r.timestamp), "12345");
