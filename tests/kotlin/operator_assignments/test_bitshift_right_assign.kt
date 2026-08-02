// vybe-test: kotlin/operator_assignments/test_bitshift_right_assign
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 16
a = a shr 2
__check((a).toString(), "4") }
