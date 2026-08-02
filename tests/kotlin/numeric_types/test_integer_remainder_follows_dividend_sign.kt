// vybe-test: kotlin/numeric_types/test_integer_remainder_follows_dividend_sign
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((7 % 3).toString(), "1")
            __check((7 % 4).toString(), "3")
            __check((-7 % 3).toString(), "-1")
            __check((7 % -4).toString(), "3")
        }
