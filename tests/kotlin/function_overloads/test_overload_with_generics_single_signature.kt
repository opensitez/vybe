// vybe-test: kotlin/function_overloads/test_overload_with_generics_single_signature
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun <T> join(values: List<T>): Int = values.size
        fun join(values: String): Int = values.length
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((join(listOf(1, 2))).toString(), "2")
            __check((join("ab")).toString(), "2")
        }
