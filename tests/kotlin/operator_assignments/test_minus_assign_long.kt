// vybe-test: kotlin/operator_assignments/test_minus_assign_long
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a: Long = 20L
a -= 5L
__check((a).toString(), "15") }
