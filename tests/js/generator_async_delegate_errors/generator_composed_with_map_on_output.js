// vybe-test: js/generator_async_delegate_errors/generator_composed_with_map_on_output
// origin: languages/js/tests/js/test_generator_async_delegate_errors.rs

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

function* nums(){yield 1;yield 2;} const mapped=[...nums()].map(x=>x*10); __check(__line(mapped.join(",")), "10,20");
