// vybe-test: kotlin/imports/test_import_kotlin_math_sqrt
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.sqrt
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sqrt(16.0).toInt()).toString(), "4")
        }
