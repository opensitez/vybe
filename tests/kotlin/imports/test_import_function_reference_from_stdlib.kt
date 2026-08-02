// vybe-test: kotlin/imports/test_import_function_reference_from_stdlib
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.abs
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = ::abs
            __check((f(-10)).toString(), "10")
        }
