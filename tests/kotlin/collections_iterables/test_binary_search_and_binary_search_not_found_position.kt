// vybe-test: kotlin/collections_iterables/test_binary_search_and_binary_search_not_found_position
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 3, 5, 7, 9)
            __check((nums.binarySearch(5)).toString(), "2")
            __check((nums.binarySearch(6)).toString(), "-3")
        }
