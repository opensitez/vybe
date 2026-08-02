// vybe-test: kotlin/imports/test_import_array_functions
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.collections.maxOrNull
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((listOf(1, 5, 3).maxOrNull()).toString(), "5")
        }
