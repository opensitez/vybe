// vybe-test: kotlin/generic_constraints/test_generic_constraints_list_of_chars
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : CharSequence> concat(values: List<T>): String = values.joinToString(":")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((concat(listOf("a", "b"))).toString(), "a:b")
        }
