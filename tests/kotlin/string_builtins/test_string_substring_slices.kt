// vybe-test: kotlin/string_builtins/test_string_substring_slices
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "kotlin-lang"
            __check((text.substring(0, 6)).toString(), "kotlin")
            __check((text.substring(7)).toString(), "lang")
        }
