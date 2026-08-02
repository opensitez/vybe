// vybe-test: kotlin/kotlin_package_aliases/test_multiple_aliases_for_same_symbol_family
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.collections.joinToString as joinA
        import kotlin.collections.joinToString as joinB

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = listOf(1, 2).joinA(",")
            __check((text).toString(), "1,2")
            __check((joinB(listOf(3, 4), "+")).toString(), "3+4")
        }
