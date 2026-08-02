// vybe-test: kotlin/operator_assignments/test_assign_with_return_and_print
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var a = 1
        __check((a).toString(), "1")
        a = a + 4
        __check((a).toString(), "5")
    }
