// vybe-test: kotlin/imports/test_import_type_alias_name_collision_with_local
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.collections.List
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: List<Int> = listOf(1, 2, 3)
            __check((values.size).toString(), "3")
        }
