// vybe-test: kotlin/kotlin_number_parsing/test_to_int_or_null_handles_invalid_input
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_parsing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("x".toIntOrNull()).toString(), "null")
            __check(("".toIntOrNull()).toString(), "null")
        }
