// vybe-test: js/string_processing_patterns/snake_to_camel_case
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

function toCamelCase(str) {
    return str.replace(/_(\w)/g, (_, c) => c.toUpperCase());
}
__check(__line(toCamelCase("hello_world")), "helloWorld");
__check(__line(toCamelCase("some_variable_name")), "someVariableName");
