// vybe-test: js/generator_protocol_advanced/generator_as_state_machine
// origin: languages/js/tests/js/test_generator_protocol_advanced.rs

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

function* trafficLight() {
    while (true) {
        yield "red";
        yield "green";
        yield "yellow";
    }
}
const light = trafficLight();
const states = [];
for (let i = 0; i < 5; i++) states.push(light.next().value);
console.log(states.join(","));
