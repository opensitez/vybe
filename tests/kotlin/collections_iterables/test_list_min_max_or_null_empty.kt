// vybe-test: kotlin/collections_iterables/test_list_min_max_or_null_empty
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(5, 2, 9, 1)
            __check((nums.minOrNull()).toString(), "1")
            __check((nums.maxOrNull()).toString(), "9")
            __check((listOf<Int>().minOrNull() ?: -1).toString(), "-1")
        }
