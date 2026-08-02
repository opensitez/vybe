// vybe-test: kotlin/import_aliases/test_import_alias_star_kept_local_shadowing
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.collections.*

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mutableList = mutableListOf("a", "b")
            __check((mutableList.joinToString("/")).toString(), "a/b")
            __check((listOf(1, 2).size).toString(), "2")
        }
