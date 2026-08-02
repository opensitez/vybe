// vybe-test: kotlin/import_aliases/test_import_alias_chain_of_calls
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.math.max as takeMax
        import kotlin.math.min as takeMin

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((takeMax(3, 9)).toString(), "9")
            __check((takeMin(3, 9)).toString(), "3")
        }
