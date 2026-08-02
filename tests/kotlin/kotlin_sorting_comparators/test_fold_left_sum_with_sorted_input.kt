// vybe-test: kotlin/kotlin_sorting_comparators/test_fold_left_sum_with_sorted_input
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(9, 1, 5, 3).sorted()
            val total = values.fold(0) { acc, n -> acc + n }
            __check((values.joinToString(",")).toString(), "1,3,5,9")
            __check((total).toString(), "18")
        }
