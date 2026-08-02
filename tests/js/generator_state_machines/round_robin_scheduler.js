// vybe-test: js/generator_state_machines/round_robin_scheduler
// origin: languages/js/tests/js/test_generator_state_machines.rs

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

function* roundRobin(tasks) {
    const generators = tasks.map(t => t());
    while (generators.length > 0) {
        for (let i = generators.length - 1; i >= 0; i--) {
            const result = generators[i].next();
            if (result.done) generators.splice(i, 1);
            else yield result.value;
        }
    }
}
function* task(name, steps) {
    for (let i = 0; i < steps; i++) yield `${name}:${i}`;
}
const log = [...roundRobin([
    () => task("A", 2),
    () => task("B", 2),
])];
console.log(log.join(","));
