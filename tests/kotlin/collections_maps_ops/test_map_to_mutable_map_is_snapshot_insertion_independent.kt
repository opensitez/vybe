// vybe-test: kotlin/collections_maps_ops/test_map_to_mutable_map_is_snapshot_insertion_independent
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = mapOf("x" to 1, "y" to 2)
            val copied = base.toMutableMap()
            copied["z"] = 3
            __check((base["z"]).toString(), "null")
            __check((copied["z"]).toString(), "3")
            __check((copied.size).toString(), "3")
            __check((base.size).toString(), "2")
        }
