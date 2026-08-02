// vybe-test: kotlin/operator_assignments/test_compound_double_rem
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 10.0
a = a % 3.0
__check((a).toString(), "1.0") }
