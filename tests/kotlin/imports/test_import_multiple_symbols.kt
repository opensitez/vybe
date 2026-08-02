// vybe-test: kotlin/imports/test_import_multiple_symbols
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.abs
        import kotlin.math.roundToInt
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((abs(-2)).toString(), "2")
            __check((2.8.roundToInt()).toString(), "3")
        }
