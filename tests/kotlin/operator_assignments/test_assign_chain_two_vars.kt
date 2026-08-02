// vybe-test: kotlin/operator_assignments/test_assign_chain_two_vars
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var a = 1
        var b = 2
        a += b
        b += a
        __check((a).toString(), "3")
        __check((b).toString(), "5")
    }
