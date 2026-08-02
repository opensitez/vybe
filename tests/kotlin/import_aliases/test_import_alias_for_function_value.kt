// vybe-test: kotlin/import_aliases/test_import_alias_for_function_value
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.math.max as pickMax

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pickMax(3, 10)).toString(), "10")
            __check((pickMax(-1, 2)).toString(), "2")
        }
