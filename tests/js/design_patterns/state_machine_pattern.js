// vybe-test: js/design_patterns/state_machine_pattern
// origin: languages/js/tests/js/test_design_patterns.rs

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

class TrafficLight {
    #state = "red";
    next() {
        const transitions = { red: "green", green: "yellow", yellow: "red" };
        this.#state = transitions[this.#state];
        return this.#state;
    }
    get state() { return this.#state; }
}
const light = new TrafficLight();
__check(__line(light.state), "red");
__check(__line(light.next()), "green");
__check(__line(light.next()), "yellow");
__check(__line(light.next()), "red");
