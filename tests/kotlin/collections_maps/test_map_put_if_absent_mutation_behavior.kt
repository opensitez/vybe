// vybe-test: kotlin/collections_maps/test_map_put_if_absent_mutation_behavior
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counts = mutableMapOf("a" to 1)
            __check((counts.put("a", 9)).toString(), "1")
            __check((counts.putIfAbsent("a", 11)).toString(), "1")
            __check((counts.putIfAbsent("b", 2)).toString(), "null")
            __check((counts["a"]).toString(), "9")
            __check((counts["b"]).toString(), "2")
        }
