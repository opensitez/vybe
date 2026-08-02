// vybe-test: kotlin/imports/test_import_multiple_aliases_same_base
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.sqrt as sq
        import kotlin.math.sqrt as squareRoot
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sq(9.0).toInt()).toString(), "3")
            __check((squareRoot(16.0).toInt()).toString(), "4")
        }
