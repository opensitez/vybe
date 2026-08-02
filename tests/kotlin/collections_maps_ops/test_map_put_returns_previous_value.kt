// vybe-test: kotlin/collections_maps_ops/test_map_put_returns_previous_value
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            __check((map.put("a", 5)).toString(), "1")
            __check((map.put("c", 3)).toString(), "null")
            __check((map["a"] + map["c"]).toString(), "8")
        }
