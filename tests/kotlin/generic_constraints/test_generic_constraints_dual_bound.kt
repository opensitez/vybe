// vybe-test: kotlin/generic_constraints/test_generic_constraints_dual_bound
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

interface Named { val name: String }
        fun <T> label(v: T): String where T : Number, T : Comparable<T> {
            return v.toString()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(3)).toString(), "3")
        }
