// vybe-test: kotlin/literals/test_integer_literals_decimal_and_negative
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0).toString(), "0")
            __check((123).toString(), "123")
            __check((-456).toString(), "-456")
        }
