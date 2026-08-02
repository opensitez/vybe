// vybe-test: kotlin/import_aliases/test_import_alias_isolated_scopes
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.math.abs as absValue

        fun score(v: Int): Int {
            val absValue = { x: Int -> x * x }
            return absValue(v)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(-3)).toString(), "9")
            __check((absValue(-3)).toString(), "3")
        }
