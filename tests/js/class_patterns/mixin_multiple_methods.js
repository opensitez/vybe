// vybe-test: js/class_patterns/mixin_multiple_methods
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

function Timestamped(Base) {
    return class extends Base {
        getTimestamp() { return "2024-01-01"; }
    };
}
function Tagged(Base) {
    return class extends Base {
        setTag(tag) { this._tag = tag; }
        getTag() { return this._tag; }
    };
}
class Item {}
class TaggedItem extends Tagged(Timestamped(Item)) {}
let item = new TaggedItem();
item.setTag("important");
__check(__line(item.getTag()), "important");
__check(__line(item.getTimestamp()), "2024-01-01");
