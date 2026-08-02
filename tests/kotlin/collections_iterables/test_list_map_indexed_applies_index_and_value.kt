// vybe-test: kotlin/collections_iterables/test_list_map_indexed_applies_index_and_value
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(10, 20, 30)
            val withIndex = nums.mapIndexed { index, value -> value + index }
            __check((withIndex.joinToString(",")).toString(), "10,21,32")
        }
