// vybe-test: kotlin/collections_maps/test_map_get_or_put_default_does_not_override_existing
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counters = mutableMapOf("x" to 1)
            __check((counters.getOrPut("x") { 9 }).toString(), "1")
            __check((counters["x"]).toString(), "1")
            __check((counters.getOrPut("y") { 4 }).toString(), "4")
            __check((counters["y"]).toString(), "4")
            __check((counters.size).toString(), "2")
        }
