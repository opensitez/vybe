// vybe-test: kotlin/operator_assignments/test_unsigned_shift_right_not
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = -8
a = a ushr 1
__check((a).toString(), "2147483644") }
