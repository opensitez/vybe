// vybe-test: kotlin/collections_maps_ops/test_map_to_map_snapshot_is_not_reactive
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mutableMapOf("a" to 1)
            val snap = source.toMap()
            source["a"] = 4
            source["b"] = 2
            __check((snap["a"]).toString(), "1")
            __check((snap.size).toString(), "1")
            __check((source.size).toString(), "2")
        }
