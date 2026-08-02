// vybe-test: js/regex_named_groups/matchall_with_named_capture_groups
// origin: languages/js/tests/js/test_regex_named_groups.rs

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

const re = /(?<key>\w+)=(?<val>\w+)/g;
const results = [];
for (const m of "a=1 b=2 c=3".matchAll(re)) {
    results.push(m.groups.key + ":" + m.groups.val);
}
console.log(results.join(","));
