// vybe-test: kotlin/kotlin_char_apis/test_char_to_digit_and_back
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('9'.digitToInt()).toString(), "9")
            __check(('a'.digitToIntOrNull()).toString(), "null")
            __check(('a'.digitToInt(16)).toString(), "10")
            __check(('f'.digitToIntOrNull(16)).toString(), "15")
        }
