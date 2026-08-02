// vybe-test: kotlin/collections_maps_ops/test_map_remove_returns_value_or_null
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            __check((map.remove("a")).toString(), "1")
            __check((map.remove("missing")).toString(), "null")
            __check((map.size).toString(), "1")
        }
