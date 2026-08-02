// vybe-test: kotlin/reified_generics/test_reified_list_of_types
// origin: languages/kotlin/tests/kotlin/test_reified_generics.rs

inline fun <reified T> describe(values: List<T>): String = values::class.simpleName ?: ""

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(listOf(1, 2, 3))).toString(), "ArrayList")
            __check((describe(listOf("a", "b"))).toString(), "ArrayList")
        }
