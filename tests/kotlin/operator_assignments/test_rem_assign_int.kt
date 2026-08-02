// vybe-test: kotlin/operator_assignments/test_rem_assign_int
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 17
a %= 5
__check((a).toString(), "2") }
