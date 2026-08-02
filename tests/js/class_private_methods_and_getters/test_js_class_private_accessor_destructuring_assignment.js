// vybe-test: js/class_private_methods_and_getters/test_js_class_private_accessor_destructuring_assignment
// origin: languages/js/tests/js/test_js_class_private_methods_and_getters.rs

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

class Coords {
    #x = 10; #y = 20;
    get #pair() { return [this.#x, this.#y]; }
    set #pair([x, y]) { this.#x = x; this.#y = y; }

    update(x, y) {
        this.#pair = [x, y];
    }
    read() {
        const [x, y] = this.#pair;
        return `${x}:${y}`;
    }
}
const c = new Coords();
c.update(100, 200);
__check(__line(c.read()), "100:200");
