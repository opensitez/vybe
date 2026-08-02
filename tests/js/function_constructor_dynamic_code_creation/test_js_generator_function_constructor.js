// vybe-test: js/function_constructor_dynamic_code_creation/test_js_generator_function_constructor
// origin: languages/js/tests/js/test_js_function_constructor_dynamic_code_creation.rs

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

const GeneratorFunction = Object.getPrototypeOf(function*(){}).constructor;
const gen = new GeneratorFunction("a", "yield a * 10; yield a * 20;");
const g = gen(5);
__check(__line(`${g.next().value}:${g.next().value}`), "50:100");
