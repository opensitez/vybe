// vybe-test: kotlin/imports/test_import_java_math_namespace
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import java.lang.Math
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Math.max(2, 8)).toString(), "8")
        }
