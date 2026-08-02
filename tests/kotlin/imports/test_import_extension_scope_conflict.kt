// vybe-test: kotlin/imports/test_import_extension_scope_conflict
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.text.uppercase
        import kotlin.text.lowercase
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "Ab"
            __check((text.uppercase()).toString(), "AB")
            __check((text.lowercase()).toString(), "ab")
        }
