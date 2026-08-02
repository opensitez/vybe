// vybe-test: kotlin/collections_maps/test_map_put_all_from_pairs_sequence
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mutableMapOf("a" to 1)
            source.putAll(listOf("a" to 9, "b" to 2).asSequence())
            __check((source["a"]).toString(), "9")
            __check((source["b"]).toString(), "2")
            __check((source.size).toString(), "2")
        }
