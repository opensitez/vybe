// vybe-test: kotlin/collections_maps_ops/test_map_keys_view_removal_mutates_backing_map
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            val keys = map.keys
            keys.remove("b")
            __check((map.size).toString(), "2")
            __check((map.containsKey("b")).toString(), "false")
            __check((map["a"] + map["c"]).toString(), "4")
        }
