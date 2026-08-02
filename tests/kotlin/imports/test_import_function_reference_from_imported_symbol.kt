// vybe-test: kotlin/imports/test_import_function_reference_from_imported_symbol
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.absoluteValue
        fun norm(v: Int) = v.absoluteValue
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((norm(-12)).toString(), "12")
        }
