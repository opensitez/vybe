// vybe-test: kotlin/collections_iterables/test_grouping_by_each_count
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 1, 2, 2, 2, 3)
            val counts = nums.groupingBy { it }.eachCount()
            __check((counts[1]).toString(), "2")
            __check((counts[2]).toString(), "3")
            __check((counts[3]).toString(), "1")
        }
