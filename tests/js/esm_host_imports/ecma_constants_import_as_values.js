// vybe-test: js/esm_host_imports/ecma_constants_import_as_values
// origin: languages/js/tests/js/test_esm_host_imports.rs

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

import { PI, E } from "ecma:math";
import { MAX_SAFE_INTEGER, NaN as NumberNaN } from "ecma:number";
__check(__line(typeof PI === "number"), "true");
__check(__line(PI > 3), "true");
__check(__line(typeof E === "number"), "true");
__check(__line(MAX_SAFE_INTEGER > 1000), "true");
__check(__line(Number.isNaN(NumberNaN)), "true");
