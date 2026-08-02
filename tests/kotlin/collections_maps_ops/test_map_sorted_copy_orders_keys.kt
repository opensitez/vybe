// vybe-test: kotlin/collections_maps_ops/test_map_sorted_copy_orders_keys
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("c" to 3, "a" to 1, "b" to 2)
            val sorted = map.toSortedMap()
            __check((sorted.keys.joinToString(",")).toString(), "a,b,c")
            __check((sorted.values.joinToString(",")).toString(), "1,2,3")
        }
