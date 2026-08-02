// vybe-test: kotlin/imports/test_import_sequence_extension_call
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.collections.reduce
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sum = listOf(1, 2, 3).reduce { a, b -> a + b }
            __check((sum).toString(), "6")
        }
