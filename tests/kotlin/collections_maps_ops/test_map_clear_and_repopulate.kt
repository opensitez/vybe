// vybe-test: kotlin/collections_maps_ops/test_map_clear_and_repopulate
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            map.clear()
            __check((map.isEmpty()).toString(), "true")
            map["z"] = 10
            __check((map.size).toString(), "1")
            __check((map["z"]).toString(), "10")
        }
