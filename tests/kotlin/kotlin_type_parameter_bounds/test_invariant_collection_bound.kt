// vybe-test: kotlin/kotlin_type_parameter_bounds/test_invariant_collection_bound
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

fun <T> firstOf(list: List<T>): T {
            return list[0]
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((firstOf(listOf(1, 2, 3))).toString(), "1")
            __check((firstOf(listOf("a", "b"))).toString(), "a")
        }
