// vybe-test: kotlin/generic_constraints/test_generic_constraints_charsequence_has_prefix
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : CharSequence> begins(v: T, prefix: String): Boolean = v.startsWith(prefix)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((begins("hello", "he")).toString(), "true")
            __check((begins("hello", "x")).toString(), "false")
        }
