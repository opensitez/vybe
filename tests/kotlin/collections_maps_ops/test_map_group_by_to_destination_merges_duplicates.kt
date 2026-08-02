// vybe-test: kotlin/collections_maps_ops/test_map_group_by_to_destination_merges_duplicates
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("a", "bb", "cc", "ddd", "eee")
            val groups = mutableMapOf<Int, MutableList<String>>()
            words.groupByTo(groups, { it.length })
            __check((groups[1]?.size).toString(), "1")
            __check((groups[2]?.size).toString(), "2")
            __check((groups[3]?.size).toString(), "2")
        }
