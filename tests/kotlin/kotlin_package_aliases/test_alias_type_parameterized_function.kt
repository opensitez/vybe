// vybe-test: kotlin/kotlin_package_aliases/test_alias_type_parameterized_function
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.collections.sortedBy as bySort

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("z", "aa", "bbb")
            val sorted = values.bySort { it.length }
            __check((sorted.joinToString(",")).toString(), "z,aa,bbb")
        }
