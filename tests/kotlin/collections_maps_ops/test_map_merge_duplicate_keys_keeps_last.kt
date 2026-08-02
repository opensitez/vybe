// vybe-test: kotlin/collections_maps_ops/test_map_merge_duplicate_keys_keeps_last
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pairs = listOf("a" to 1, "b" to 2, "a" to 3)
            val map = pairs.toMap()
            __check((map["a"]).toString(), "3")
            __check((map.size).toString(), "2")
            __check((map["b"]).toString(), "2")
        }
