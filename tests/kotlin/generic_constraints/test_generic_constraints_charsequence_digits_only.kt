// vybe-test: kotlin/generic_constraints/test_generic_constraints_charsequence_digits_only
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : CharSequence> isDigits(v: T): Boolean = v.all { it.isDigit() }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isDigits("1234")).toString(), "true")
            __check((isDigits("12a4")).toString(), "false")
        }
