// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_with_reverse_numeric_comparator
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 3, 2, 5, 4)
            val out = values.sortedWith(compareByDescending { it })
            __check((out.joinToString(",")).toString(), "5,4,3,2,1")
        }
