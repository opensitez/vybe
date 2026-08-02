// vybe-test: kotlin/operator_assignments/test_times_assign
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 4
a *= 3
__check((a).toString(), "12") }
