// vybe-test: kotlin/collections_iterables/test_list_zip_with_next_produces_adjacent_pairs
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            val pairs = nums.zipWithNext { a, b -> "$a:$b" }
            __check((pairs.joinToString(",")).toString(), "1:2,2:3,3:4")
        }
