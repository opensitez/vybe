// vybe-test: kotlin/imports/test_import_static_members_on_string
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.text.appendLine
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sb = StringBuilder()
            sb.appendLine("a")
            sb.appendLine("b")
            __check((sb.toString().trim()).toString(), "a\nb")
        }
