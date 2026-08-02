// vybe-test: kotlin/collections_maps_ops/test_map_get_or_put_existing_preserves_value
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1)
            val first = map.getOrPut("a") { 9 }
            val second = map.getOrPut("b") { 9 }
            __check((first).toString(), "1")
            __check((second).toString(), "9")
            __check((map["a"]).toString(), "1")
            __check((map["b"]).toString(), "9")
        }
