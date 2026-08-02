// vybe-test: js/wasi/test_fs_write_read
// origin: languages/js/tests/js/js_wasi_test.rs

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

fs.writeFile("/tmp/vybe_test_fs.txt", "hello vybe");
        let content = fs.readFile("/tmp/vybe_test_fs.txt");
        __check(__line(content), "hello vybe");
        fs.remove("/tmp/vybe_test_fs.txt");
