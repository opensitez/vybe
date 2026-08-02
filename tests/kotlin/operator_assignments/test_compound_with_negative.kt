// vybe-test: kotlin/operator_assignments/test_compound_with_negative
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 10
a += -3
__check((a).toString(), "7") }
