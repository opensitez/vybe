// vybe-test: js/arrow_functions/arrow_bind_binds_args_not_this
// origin: languages/js/tests/js/test_arrow_functions.rs

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

const add=(a,b)=>a+b; const inc=add.bind({ignored:1},1); __check(__line(inc(5)), "6");
