// vybe-test: kotlin/collections_maps_ops/test_map_put_if_absent
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1)
            val existing = map.putIfAbsent("a", 9)
            val added = map.putIfAbsent("b", 2)
            __check((existing).toString(), "1")
            __check((added).toString(), "null")
            __check((map["a"]).toString(), "1")
            __check((map["b"]).toString(), "2")
        }
