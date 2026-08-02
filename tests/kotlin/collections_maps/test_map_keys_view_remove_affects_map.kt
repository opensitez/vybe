// vybe-test: kotlin/collections_maps/test_map_keys_view_remove_affects_map
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            val keys = source.keys
            keys.remove("b")
            __check((source.size).toString(), "2")
            __check((source.containsKey("b")).toString(), "false")
            __check((source.containsKey("c")).toString(), "true")
        }
