// vybe-test: kotlin/generic_constraints/test_generic_constraints_pair_compare
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Comparable<T>> greater(a: Pair<T, T>): T = if (a.first > a.second) a.first else a.second
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((greater(Pair(2, 9))).toString(), "9")
            __check((greater(Pair("x", "y"))).toString(), "y")
        }
