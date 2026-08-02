// vybe-test: kotlin/collections_maps_ops/test_map_update_and_remove_cycle
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("x" to 1)
            map["x"] = 2
            map.remove("x")
            map["x"] = 5
            __check((map["x"]).toString(), "5")
            __check((map.size).toString(), "1")
        }
