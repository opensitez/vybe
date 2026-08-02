// vybe-test: kotlin/imports/test_import_kotlin_math_abs
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.abs
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((abs(-5)).toString(), "5")
        }
