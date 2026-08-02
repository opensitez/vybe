// vybe-test: kotlin/generic_constraints/test_generic_constraints_mixed_constraints
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T> describe(v: T): String where T : Number, T : Comparable<T> {
            return if (v.toInt() > 4) "big" else "small"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(9)).toString(), "big")
        }
