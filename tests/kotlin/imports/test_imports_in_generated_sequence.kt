// vybe-test: kotlin/imports/test_imports_in_generated_sequence
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.max
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = generateSequence(0) { it + 1 }.take(4).toList()
            __check((values.maxOrNull()).toString(), "3")
        }
