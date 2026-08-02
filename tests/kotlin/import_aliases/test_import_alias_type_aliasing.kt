// vybe-test: kotlin/import_aliases/test_import_alias_type_aliasing
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.collections.HashMap as MapAlias
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = MapAlias<String, Int>()
            map["a"] = 1
            __check((map["a"]).toString(), "1")
        }
