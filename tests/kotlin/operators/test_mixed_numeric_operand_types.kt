// vybe-test: kotlin/operators/test_mixed_numeric_operand_types
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 + 2.0).toString(), "3")
            __check((5.5 - 2).toString(), "3.5")
            __check((8 / 4.0).toString(), "2")
            __check((8L / 3L).toString(), "2")
        }
