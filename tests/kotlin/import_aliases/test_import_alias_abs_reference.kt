// vybe-test: kotlin/import_aliases/test_import_alias_abs_reference
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.math.absoluteValue as absAlias

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((absAlias(-5)).toString(), "5")
            __check((absAlias(6)).toString(), "6")
        }
