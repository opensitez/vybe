// vybe-test: kotlin/collections_iterables/test_list_fold_left_and_reduce_right
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            val sum = nums.fold(0) { acc, v -> acc + v }
            val diff = nums.reduceRight { a, b -> a - b }
            __check((sum).toString(), "10")
            __check((diff).toString(), "-2")
        }
