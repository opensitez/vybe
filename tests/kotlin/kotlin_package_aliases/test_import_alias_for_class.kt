// vybe-test: kotlin/kotlin_package_aliases/test_import_alias_for_class
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.collections.HashMap as KMap

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map: KMap<String, Int> = KMap()
            map["a"] = 1
            __check((map["a"]).toString(), "1")
        }
