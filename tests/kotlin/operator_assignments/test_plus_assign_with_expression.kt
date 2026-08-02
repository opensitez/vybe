// vybe-test: kotlin/operator_assignments/test_plus_assign_with_expression
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 1
val b = 2
a += b * 2
__check((a).toString(), "5") }
