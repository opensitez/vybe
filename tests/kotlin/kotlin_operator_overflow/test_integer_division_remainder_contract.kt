// vybe-test: kotlin/kotlin_operator_overflow/test_integer_division_remainder_contract
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((7 / 2).toString(), "3")
            __check((7 % 2).toString(), "1")
            __check(((-7) / 2).toString(), "-3")
            __check(((-7) % 2).toString(), "-1")
        }
