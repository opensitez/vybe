// vybe-test: js/string_processing_deep/diff_two_strings
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

function changedWords(a, b) {
    const wa = a.split(" "), wb = b.split(" ");
    const changes = [];
    const maxLen = Math.max(wa.length, wb.length);
    for (let i = 0; i < maxLen; i++) {
        if (wa[i] !== wb[i]) changes.push(i);
    }
    return changes;
}
console.log(changedWords("hello world foo", "hello bar foo").join(","));
console.log(changedWords("a b c", "a b c").length);
