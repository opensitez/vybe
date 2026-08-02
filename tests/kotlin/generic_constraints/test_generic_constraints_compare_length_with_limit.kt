// vybe-test: kotlin/generic_constraints/test_generic_constraints_compare_length_with_limit
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : CharSequence> exceeds(v: T, limit: Int): Boolean = v.length > limit
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((exceeds("abc", 2)).toString(), "true")
            __check((exceeds("a", 3)).toString(), "false")
        }
