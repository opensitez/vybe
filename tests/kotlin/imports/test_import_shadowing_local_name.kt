// vybe-test: kotlin/imports/test_import_shadowing_local_name
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.abs as absoluteValue
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val abs = 99
            __check((absoluteValue(-7)).toString(), "7")
            __check((abs).toString(), "99")
        }
