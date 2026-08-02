// vybe-test: kotlin/kotlin_package_aliases/test_import_alias_for_nested_type_reference
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.collections.List as KList

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: KList<Int> = listOf(4, 5, 6)
            __check((values.sum()).toString(), "15")
        }
