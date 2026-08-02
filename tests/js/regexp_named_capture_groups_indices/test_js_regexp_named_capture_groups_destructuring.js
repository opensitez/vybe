// vybe-test: js/regexp_named_capture_groups_indices/test_js_regexp_named_capture_groups_destructuring
// origin: languages/js/tests/js/test_js_regexp_named_capture_groups_indices.rs

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

const re = /(?<title>Mr\.|Mrs\.|Dr\.) (?<name>\w+)/;
const { groups: { title, name } } = re.exec("Dr. Watson");
__check(__line(`${title} ${name}`), "Dr. Watson");
