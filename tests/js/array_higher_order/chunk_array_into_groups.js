// vybe-test: js/array_higher_order/chunk_array_into_groups
// origin: languages/js/tests/js/test_array_higher_order.rs

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

function chunk(arr, size) {
    const result = [];
    for (let i = 0; i < arr.length; i += size) {
        result.push(arr.slice(i, i + size));
    }
    return result;
}
const chunks = chunk([1, 2, 3, 4, 5, 6, 7], 3);
console.log(chunks.length);
console.log(chunks[0].join(","));
console.log(chunks[1].join(","));
console.log(chunks[2].join(","));
