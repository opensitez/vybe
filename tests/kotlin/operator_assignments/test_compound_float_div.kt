// vybe-test: kotlin/operator_assignments/test_compound_float_div
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 10.0f
a /= 4f
__check((a).toString(), "2.5") }
