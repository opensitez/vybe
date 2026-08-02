// vybe-test: kotlin/imports/test_import_static_extension_member
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.text.capitalize
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("kotlin".capitalize()).toString(), "Kotlin")
        }
