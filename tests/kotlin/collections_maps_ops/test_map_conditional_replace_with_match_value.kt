// vybe-test: kotlin/collections_maps_ops/test_map_conditional_replace_with_match_value
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            __check((map.replace("a", 1, 9)).toString(), "true")
            __check((map.replace("a", 2, 10)).toString(), "false")
            __check((map["a"]).toString(), "9")
        }
