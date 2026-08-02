// vybe-test: kotlin/imports/test_import_nested_package_style
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.system.exitProcess
        fun status(v: Int): String = if (v > 0) "ok" else "bad"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((status(3)).toString(), "ok")
            __check((status(0)).toString(), "bad")
        }
