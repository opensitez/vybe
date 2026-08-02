// vybe-test: js/string_processing_deep/string_tokenizer
// origin: languages/js/tests/js/test_string_processing_deep.rs

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

function* tokenize(str) {
    const re = /(\d+\.?\d*)|([a-zA-Z_]\w*)|([+\-*\/()=])/g;
    let m;
    while ((m = re.exec(str)) !== null) {
        if (m[1]) yield { type: "number", value: m[1] };
        else if (m[2]) yield { type: "ident", value: m[2] };
        else yield { type: "op", value: m[3] };
    }
}
const tokens = [...tokenize("x = 3.14 + y")];
console.log(tokens.length);
console.log(tokens[0].type + ":" + tokens[0].value);
console.log(tokens[2].type + ":" + tokens[2].value);
