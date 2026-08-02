// vybe-test: kotlin/imports/test_import_array_of_functions
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.abs
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ops: Array<(Int) -> Int> = arrayOf({ abs(it) }, { it * 2 })
            __check((ops[0](-3)).toString(), "3")
            __check((ops[1](4)).toString(), "8")
        }
