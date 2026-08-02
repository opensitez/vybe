// vybe-test: kotlin/generic_constraints/test_generic_constraints_where_bound_charsequence
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T> first(v: T): Char where T : CharSequence {
            return v.first()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((first("xy")).toString(), "x")
        }
