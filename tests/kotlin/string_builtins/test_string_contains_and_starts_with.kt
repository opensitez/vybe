// vybe-test: kotlin/string_builtins/test_string_contains_and_starts_with
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "language"
            __check((text.startsWith("lang")).toString(), "true")
            __check((text.contains("gua")).toString(), "true")
            __check((text.endsWith("age")).toString(), "true")
        }
