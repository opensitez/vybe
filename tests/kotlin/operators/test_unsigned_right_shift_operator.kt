// vybe-test: kotlin/operators/test_unsigned_right_shift_operator
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((-1 ushr 1).toString(), "2147483647")
            __check((-32 ushr 3).toString(), "536870911")
            __check((16 ushr 1).toString(), "8")
        }
