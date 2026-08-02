// vybe-test: kotlin/operator_assignments/test_assign_ternary_like
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var x = 2
        val y = if (x < 0) x - 1 else x + 1
        __check((y).toString(), "3")
    }
