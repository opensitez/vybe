// vybe-test: kotlin/generic_constraints/test_generic_constraints_pair_of_numbers
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> totalPair(a: T, b: T): Double = a.toDouble() + b.toDouble()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((totalPair(1, 2)).toString(), "3.0")
            __check((totalPair(1.5, 2.25)).toString(), "3.75")
        }
