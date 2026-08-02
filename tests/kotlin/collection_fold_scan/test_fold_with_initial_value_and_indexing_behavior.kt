// vybe-test: kotlin/collection_fold_scan/test_fold_with_initial_value_and_indexing_behavior
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3)
            val total = nums.fold(10) { acc, n -> acc + n }
            val product = nums.fold(1) { acc, n -> acc * n }
            __check((total).toString(), "16")
            __check((product).toString(), "6")
        }
