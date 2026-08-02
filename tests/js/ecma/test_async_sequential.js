// vybe-test: js/ecma/test_async_sequential
// origin: languages/js/tests/js/js_ecma_test.rs

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

async function step1() { return 10; }
        async function step2(x) { return x * 2; }
        async function step3(x) { return x + 5; }

        let a = await step1();
        let b = await step2(a);
        let c = await step3(b);
        console.log(c);
