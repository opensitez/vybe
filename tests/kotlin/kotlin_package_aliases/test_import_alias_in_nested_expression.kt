// vybe-test: kotlin/kotlin_package_aliases/test_import_alias_in_nested_expression
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.collections.map as asMap

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = listOf(1, 2, 3).asMap { it * it }
            __check((out.joinToString(",")).toString(), "1,4,9")
        }
