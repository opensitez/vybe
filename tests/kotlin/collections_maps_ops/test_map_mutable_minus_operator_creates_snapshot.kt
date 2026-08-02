// vybe-test: kotlin/collections_maps_ops/test_map_mutable_minus_operator_creates_snapshot
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            val reduced = source - "b"
            source["d"] = 4
            __check((reduced.containsKey("d")).toString(), "false")
            __check((source["d"]).toString(), "4")
            __check((reduced.size).toString(), "2")
        }
