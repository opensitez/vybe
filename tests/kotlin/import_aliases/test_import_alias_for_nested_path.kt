// vybe-test: kotlin/import_aliases/test_import_alias_for_nested_path
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.collections.ArrayList as KotlinIntList

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: KotlinIntList<Int> = KotlinIntList()
            values.add(5)
            values.add(6)
            __check((values[0] + values[1]).toString(), "11")
        }
