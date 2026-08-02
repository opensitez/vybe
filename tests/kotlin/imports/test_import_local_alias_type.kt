// vybe-test: kotlin/imports/test_import_local_alias_type
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.collections.HashMap as HM
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = HM<String, Int>()
            map["x"] = 9
            __check((map["x"]).toString(), "9")
        }
