// vybe-test: kotlin/imports/test_import_alias_as
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.max as maxValue
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((maxValue(3, 7)).toString(), "7")
        }
