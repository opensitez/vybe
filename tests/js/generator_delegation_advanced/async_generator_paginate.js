// vybe-test: js/generator_delegation_advanced/async_generator_paginate
// origin: languages/js/tests/js/test_generator_delegation_advanced.rs

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

async function* paginate(data, pageSize) {
    for (let i = 0; i < data.length; i += pageSize) {
        yield data.slice(i, i + pageSize);
    }
}
async function main() {
    const pages = [];
    for await (const page of paginate([1,2,3,4,5,6,7], 3)) {
        pages.push(page.join(","));
    }
    console.log(pages.length);
    console.log(pages[0]);
    console.log(pages[2]);
}
main();
