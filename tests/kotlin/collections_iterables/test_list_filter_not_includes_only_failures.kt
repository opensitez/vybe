// vybe-test: kotlin/collections_iterables/test_list_filter_not_includes_only_failures
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            val odds = nums.filterNot { it % 2 == 0 }
            __check((odds.joinToString(",")).toString(), "1,3")
        }
