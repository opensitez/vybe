// vybe-test: kotlin/imports/test_import_invalid_alias_not_allowed_is_not_compiled
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.abs as absolute
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((absolute(-4)).toString(), "4")
        }
