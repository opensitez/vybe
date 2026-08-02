// vybe-test: kotlin/collections_iterables/test_list_map_not_empty_and_is_not_empty
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2)
            val empty = listOf<Int>()
            __check((nums.isNotEmpty()).toString(), "true")
            __check((empty.isEmpty()).toString(), "true")
        }
