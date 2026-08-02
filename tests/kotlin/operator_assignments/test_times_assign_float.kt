// vybe-test: kotlin/operator_assignments/test_times_assign_float
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 2.5f
a *= 4f
__check((a).toString(), "10.0") }
