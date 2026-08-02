// vybe-test: kotlin/collections_maps/test_mutable_map_update
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counters = mutableMapOf("a" to 1, "b" to 2)
            counters["a"] = 4
            __check((counters["a"]).toString(), "4")
            counters.remove("b")
            __check((counters.containsKey("b")).toString(), "false")
            __check((counters.size).toString(), "1")
        }
