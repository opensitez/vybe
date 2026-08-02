// vybe-test: kotlin/imports/test_import_non_ascii_path_alias
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.collections.mutableListOf as listOfAlias
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOfAlias(1, 2, 3)
            __check((values.size).toString(), "3")
        }
