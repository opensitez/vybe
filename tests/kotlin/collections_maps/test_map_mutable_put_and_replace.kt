// vybe-test: kotlin/collections_maps/test_map_mutable_put_and_replace
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counts = mutableMapOf("a" to 1)
            counts["a"] = 4
            counts["b"] = 2
            __check((counts["a"]).toString(), "4")
            __check((counts["b"]).toString(), "2")
            __check((counts.size).toString(), "2")
        }
