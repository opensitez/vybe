// vybe-test: kotlin/collections_maps/test_map_key_view_reflects_mutations
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            val keys = map.keys
            __check((keys.contains("a")).toString(), "true")
            map["c"] = 3
            __check((keys.size).toString(), "3")
            map.remove("a")
            __check((keys.contains("a")).toString(), "false")
            __check((keys.contains("c")).toString(), "true")
            map.clear()
            __check((keys.isEmpty()).toString(), "true")
        }
