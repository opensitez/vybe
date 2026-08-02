// vybe-test: kotlin/collections_maps_ops/test_map_count_and_sum_from_entries_projection
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            val sum = map.entries.map { it.value }.sum()
            val count = map.entries.filter { it.key != "b" }.count()
            __check((sum).toString(), "6")
            __check((count).toString(), "2")
        }
