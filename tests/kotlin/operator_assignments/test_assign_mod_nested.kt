// vybe-test: kotlin/operator_assignments/test_assign_mod_nested
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var x = 100
        x %= 7
        x += 1
        __check((x).toString(), "3")
    }
