// vybe-test: kotlin/operator_assignments/test_long_chain_assign
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var a: Long = 2L
        a = a + 1
        a *= 2
        a -= 1
        __check((a).toString(), "5")
    }
