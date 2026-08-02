// vybe-test: js/wasi/test_fs_stat
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

fs.writeFile("/tmp/vybe_stat_test.txt", "hello");
        let info = fs.stat("/tmp/vybe_stat_test.txt");
        __check(__line(info.isFile, info.size), "true 5");
        fs.remove("/tmp/vybe_stat_test.txt");
