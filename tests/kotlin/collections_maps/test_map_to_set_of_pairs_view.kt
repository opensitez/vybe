// vybe-test: kotlin/collections_maps/test_map_to_set_of_pairs_view
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = mapOf("a" to 1, "b" to 2)
            val pairs = data.toSet()
            __check((pairs.size).toString(), "2")
            __check((pairs.contains(Pair("a", 1))).toString(), "true")
            __check((pairs.contains(Pair("b", 3))).toString(), "false")
        }
