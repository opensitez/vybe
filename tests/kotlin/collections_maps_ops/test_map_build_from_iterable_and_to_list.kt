// vybe-test: kotlin/collections_maps_ops/test_map_build_from_iterable_and_to_list
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf("a" to 1, "b" to 2).toMap()
            val items = map.toList()
            __check((items[0].first).toString(), "a")
            __check((items[1].second).toString(), "2")
            __check((items.size).toString(), "2")
        }
