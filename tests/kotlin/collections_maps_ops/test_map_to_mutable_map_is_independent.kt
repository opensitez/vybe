// vybe-test: kotlin/collections_maps_ops/test_map_to_mutable_map_is_independent
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("a" to 1, "b" to 2)
            val copy = source.toMutableMap()
            copy["a"] = 9
            __check((source["a"]).toString(), "1")
            __check((copy["a"]).toString(), "9")
            __check((copy.size).toString(), "2")
            __check((source.size).toString(), "2")
        }
