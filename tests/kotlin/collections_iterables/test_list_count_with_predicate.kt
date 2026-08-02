// vybe-test: kotlin/collections_iterables/test_list_count_with_predicate
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            __check((nums.count { it % 2 == 1 }).toString(), "3")
        }
