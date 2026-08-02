// vybe-test: js/string_processing_deep/word_wrap
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

function wordWrap(text, width) {
    const words = text.split(" ");
    const lines = [];
    let line = "";
    for (const word of words) {
        if ((line + " " + word).trim().length <= width) {
            line = (line + " " + word).trim();
        } else {
            if (line) lines.push(line);
            line = word;
        }
    }
    if (line) lines.push(line);
    return lines;
}
const lines = wordWrap("The quick brown fox jumps over the lazy dog", 15);
console.log(lines[0]);
console.log(lines[1]);
