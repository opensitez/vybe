// vybe-test: js/control_flow_advanced/continue_inside_switch_inside_loop
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

let result = [];
for (let i = 0; i < 4; i++) {
    switch (i) {
        case 2: continue;
    }
    result.push(i);
}
console.log(result.join(","));
