// vybe-test: kotlin/generic_constraints/test_generic_constraints_restrict_to_nullable_char
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : CharSequence> safeFirst(v: T?): String = if (v == null || v.isEmpty()) "-" else v[0].toString()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((safeFirst("abc")).toString(), "a")
            __check((safeFirst(null)).toString(), "-")
        }
