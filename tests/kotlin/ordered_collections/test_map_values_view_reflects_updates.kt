// vybe-test: kotlin/ordered_collections/test_map_values_view_reflects_updates
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val values = map.values
            map["a"] = 9
            __check((values.joinToString(",")).toString(), "9,2")
        }
