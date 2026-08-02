// vybe-test: kotlin/operator_assignments/test_div_assign_double
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 9.0
a /= 3.0
__check((a).toString(), "3.0") }
