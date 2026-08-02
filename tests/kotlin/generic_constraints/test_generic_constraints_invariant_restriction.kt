// vybe-test: kotlin/generic_constraints/test_generic_constraints_invariant_restriction
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T> identity(v: T): T = v
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((identity("a")).toString(), "a")
            __check((identity(12)).toString(), "12")
        }
