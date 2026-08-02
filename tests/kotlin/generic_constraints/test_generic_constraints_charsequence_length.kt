// vybe-test: kotlin/generic_constraints/test_generic_constraints_charsequence_length
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : CharSequence> len(v: T): Int = v.length
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((len("abc")).toString(), "3")
        }
