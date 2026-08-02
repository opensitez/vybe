// vybe-test: kotlin/collections_maps/test_map_values_view_remove_by_value_affects_source
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mutableMapOf("a" to 1, "b" to 2, "c" to 2)
            val values = source.values
            __check((values.remove(2)).toString(), "true")
            __check((source.size).toString(), "2")
            __check((source["c"] ?: -1).toString(), "-1")
        }
