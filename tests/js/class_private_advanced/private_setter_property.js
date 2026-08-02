// vybe-test: js/class_private_advanced/private_setter_property
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

class Validated {
    #score = 0;
    set #safeScore(v) { this.#score = v < 0 ? 0 : v > 100 ? 100 : v; }
    setScore(v) { this.#safeScore = v; }
    getScore() { return this.#score; }
}
const obj = new Validated();
obj.setScore(150);
__check(__line(obj.getScore()), "100");
obj.setScore(-10);
__check(__line(obj.getScore()), "0");
obj.setScore(75);
__check(__line(obj.getScore()), "75");
