// vybe-test: kotlin/imports/test_import_function_and_method_combo
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.roundToInt
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1.9.roundToInt()).toString(), "2")
        }
