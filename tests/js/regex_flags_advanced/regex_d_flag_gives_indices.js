// vybe-test: js/regex_flags_advanced/regex_d_flag_gives_indices
// origin: languages/js/tests/js/test_regex_flags_advanced.rs

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

const re = /(?<name>\w+)/d;
const m = re.exec("hello world");
__check(__line(m.indices[0][0]), "0"); // start of full match
__check(__line(m.indices[0][1]), "5"); // end
__check(__line(m.indices.groups.name[0]), "0"); // group start
