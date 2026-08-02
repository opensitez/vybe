// vybe-test: kotlin/import_aliases/test_import_alias_conflict_with_local_type
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.collections.List as KotlinList
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val local: KotlinList<Int> = listOf(1, 2, 3)
            __check((local.size).toString(), "3")
        }
