// vybe-test: kotlin/generic_constraints/test_generic_constraints_charsequence_tail
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : CharSequence> tail(v: T): String = v.takeLast(1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((tail("abcd")).toString(), "d")
        }
