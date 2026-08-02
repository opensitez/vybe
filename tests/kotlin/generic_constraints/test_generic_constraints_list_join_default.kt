// vybe-test: kotlin/generic_constraints/test_generic_constraints_list_join_default
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T> joinOrDash(values: List<T>): String = if (values.isEmpty()) "-" else values.joinToString(",")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((joinOrDash(listOf<Int>())).toString(), "-")
            __check((joinOrDash(listOf("a", "b"))).toString(), "a,b")
        }
