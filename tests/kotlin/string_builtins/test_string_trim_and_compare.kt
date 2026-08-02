// vybe-test: kotlin/string_builtins/test_string_trim_and_compare
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "  Kotlin  "
            __check((text.trim()).toString(), "Kotlin")
            __check((text.trim().lowercase() == "kotlin").toString(), "true")
        }
