// vybe-test: kotlin/kotlin_sorting_comparators/test_as_reversed_views_original_mutation_reflected
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mutable = mutableListOf(1, 2, 3)
            val reversed = mutable.asReversed()
            reversed[0] = 9
            __check((mutable.joinToString(",")).toString(), "9,2,1")
            __check((reversed.joinToString(",")).toString(), "1,2,9")
        }
