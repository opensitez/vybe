// vybe-test: js/scope_prototype/tagged_template_literal
// origin: languages/js/tests/js/test_scope_prototype.rs

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

function upper(strings, ...values) {
    let result = "";
    strings.forEach((str, i) => {
        result += str;
        if (i < values.length) result += String(values[i]).toUpperCase();
    });
    return result;
}
let name = "world";
let num = 42;
console.log(upper`hello ${name} you are ${num}`);
