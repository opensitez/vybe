// vybe-test: kotlin/imports/test_import_qualified_call_after_rename
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import java.lang.StringBuilder
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = StringBuilder()
            b.append("a").append("b")
            __check((b.toString()).toString(), "ab")
        }
