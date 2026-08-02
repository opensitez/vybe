// vybe-test: kotlin/import_aliases/test_import_alias_function_alias_usage
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.math.absoluteValue as absValue
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((absValue(-11)).toString(), "11")
        }
