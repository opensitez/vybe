// vybe-test: kotlin/operators/test_arithmetic_left_associativity
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((10 - 5 - 2).toString(), "3")
            __check((10 - (5 - 2)).toString(), "7")
            __check((2 + 3 - 1 + 4).toString(), "8")
        }
