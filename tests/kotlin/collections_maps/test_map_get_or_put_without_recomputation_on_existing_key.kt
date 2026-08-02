// vybe-test: kotlin/collections_maps/test_map_get_or_put_without_recomputation_on_existing_key
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var computed = 0
            val counts = mutableMapOf("a" to 1)
            __check((counts.getOrPut("a") { computed += 1; 9 }).toString(), "1")
            __check((counts.getOrPut("b") { computed += 1; 2 }).toString(), "2")
            __check((counts["a"]).toString(), "1")
            __check((counts["b"]).toString(), "2")
            __check((computed).toString(), "1")
        }
