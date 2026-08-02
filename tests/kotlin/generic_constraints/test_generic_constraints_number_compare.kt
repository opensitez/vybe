// vybe-test: kotlin/generic_constraints/test_generic_constraints_number_compare
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Comparable<T>> top(a: T, b: T): T = if (a > b) a else b
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((top(5, 2)).toString(), "5")
            __check((top("k", "a")).toString(), "k")
        }
