// vybe-test: kotlin/imports/test_import_string_builder_from_kotlin
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.text.StringBuilder
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = StringBuilder()
            b.append("x").append("y")
            __check((b.toString()).toString(), "xy")
        }
