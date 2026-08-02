// vybe-test: kotlin/generic_constraints/test_generic_constraints_count_length_if_possible
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : CharSequence> report(v: T): Int = v.count()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((report("ab")).toString(), "2")
        }
