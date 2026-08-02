// vybe-test: kotlin/operator_assignments/test_assign_if
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var x = 1
        val y = if (x == 1) { x += 5
x } else { x }
        __check((y).toString(), "6")
    }
