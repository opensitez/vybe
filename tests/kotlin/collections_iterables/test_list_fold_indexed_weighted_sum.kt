// vybe-test: kotlin/collections_iterables/test_list_fold_indexed_weighted_sum
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            val result = nums.foldIndexed(0) { index, acc, value -> acc + index * value }
            __check((result).toString(), "20")
        }
