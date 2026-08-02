// vybe-test: kotlin/operator_precedence/test_subtraction_left_assoc
// origin: languages/kotlin/tests/kotlin/test_operator_precedence.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((20 - 5 - 3).toString(), "12")
        }
