// vybe-test: js/class_private_advanced/private_field_encapsulating_state
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

class TrafficLight {
    #state = "red";
    next() {
        if (this.#state === "red") this.#state = "green";
        else if (this.#state === "green") this.#state = "yellow";
        else this.#state = "red";
    }
    current() { return this.#state; }
}
const light = new TrafficLight();
__check(__line(light.current()), "red");
light.next();
__check(__line(light.current()), "green");
light.next();
__check(__line(light.current()), "yellow");
light.next();
__check(__line(light.current()), "red");
