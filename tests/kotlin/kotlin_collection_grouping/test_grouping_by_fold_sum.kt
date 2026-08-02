// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_fold_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 3, 2, 4, 5, 7)
            val sums = values.groupingBy { it % 2 == 0 }.fold(0) { acc, v -> acc + v }
            __check((sums[true]).toString(), "6")
            __check((sums[false]).toString(), "16")
        }
