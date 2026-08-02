// vybe-test: kotlin/kotlin_string_number_parsing/test_parse_int_success_and_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_number_parsing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("12".toInt()).toString(), "12")
            __check(("999".toIntOrNull()).toString(), "999")
            __check(("x12".toIntOrNull()).toString(), "null")
            __check(("x12".toIntOrNull()?.toString() ?: "null").toString(), "null")
        }
