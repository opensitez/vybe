// vybe-test: kotlin/operator_assignments/test_assign_boolean_not_allowed
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var ok = true
        ok = ok && false
        __check((ok).toString(), "false")
    }
