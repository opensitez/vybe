// vybe-test: js/new_globals_e2e/crypto_random_uuid_format
// origin: languages/js/tests/js/test_new_globals_e2e.rs

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

const id = crypto.randomUUID();
        // RFC 4122 v4: 36 chars total, dashes at 8/13/18/23, version 4 at idx 14
        __check(__line(id.length, id.charAt(14)), "36 4");
