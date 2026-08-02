// vybe-test: kotlin/collections_maps_ops/test_map_grouping_by_value_parity
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3, "d" to 4)
            val parity = map.entries.groupBy { it.value % 2 == 0 }
            __check((parity[true]?.size ?: 0).toString(), "2")
            __check((parity[false]?.size ?: 0).toString(), "2")
        }
