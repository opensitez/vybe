// vybe-test: js/string_processing_patterns/camel_to_snake_case
// origin: languages/js/tests/js/test_string_processing_patterns.rs

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

function toSnakeCase(str) {
    return str.replace(/([A-Z])/g, c => "_" + c.toLowerCase()).replace(/^_/, "");
}
__check(__line(toSnakeCase("helloWorld")), "hello_world");
__check(__line(toSnakeCase("camelCaseString")), "camel_case_string");
__check(__line(toSnakeCase("simpleTest")), "simple_test");
