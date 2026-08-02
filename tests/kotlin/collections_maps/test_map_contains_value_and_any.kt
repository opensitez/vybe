// vybe-test: kotlin/collections_maps/test_map_contains_value_and_any
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counters = mapOf("read" to 1, "write" to 2, "exec" to 0)
            __check((counters.containsValue(2)).toString(), "true")
            __check((counters.containsValue(3)).toString(), "false")
            __check((counters.any { it.value > 1 }).toString(), "true")
            __check((counters.all { it.key.isNotEmpty() }).toString(), "true")
        }
