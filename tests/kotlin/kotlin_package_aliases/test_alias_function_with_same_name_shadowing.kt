// vybe-test: kotlin/kotlin_package_aliases/test_alias_function_with_same_name_shadowing
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.collections.joinToString as join

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = listOf("a", "b", "c").let { join(it, "/") }
            __check((text).toString(), "a/b/c")
        }
