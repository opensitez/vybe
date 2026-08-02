// vybe-test: kotlin/kotlin_string_number_parsing/test_parse_boolean_or_default
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_number_parsing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = "x".toBooleanOrNull()
            __check((v ?: false).toString(), "false")
            val n = "x".toIntOrNull()
            __check((n ?: 0).toString(), "0")
        }
