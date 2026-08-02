// vybe-test: js/regex_named_groups/d_flag_provides_named_group_indices
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

const re = /(?<word>\w+)/d;
const m = re.exec("hello world");
__check(__line(m.indices.groups.word[0]), "0");
__check(__line(m.indices.groups.word[1]), "5");
