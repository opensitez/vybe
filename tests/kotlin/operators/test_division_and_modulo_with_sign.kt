// vybe-test: kotlin/operators/test_division_and_modulo_with_sign
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((10 / 3).toString(), "3")
            __check((10 / 3.0).toString(), "3.3333333333333335")
            __check((10 % 3).toString(), "1")
            __check((-10 % 3).toString(), "-1")
        }
