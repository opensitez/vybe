// vybe-test: kotlin/collections_iterables/test_list_map_indexed_not_null_filters_by_indexed_predicate
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            val picked = nums.mapIndexedNotNull { index, value ->
                if (index % 2 == 0) value * 10 else null
            }
            __check((picked.joinToString(",")).toString(), "10,30")
        }
