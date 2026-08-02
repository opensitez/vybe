// vybe-test: js/chaining_patterns/animation_chain
// origin: languages/js/tests/js/test_chaining_patterns.rs

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

class Animation {
    #steps = [];
    moveTo(x, y) { this.#steps.push(`move(${x},${y})`); return this; }
    scaleTo(s) { this.#steps.push(`scale(${s})`); return this; }
    rotateTo(deg) { this.#steps.push(`rotate(${deg})`); return this; }
    play() { return this.#steps.join(" -> "); }
}
const anim = new Animation()
    .moveTo(100, 200)
    .scaleTo(2)
    .rotateTo(90)
    .moveTo(0, 0);
__check(__line(anim.play()), "move(100,200) -> scale(2) -> rotate(90) -> move(0,0)");
