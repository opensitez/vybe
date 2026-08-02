// vybe-test: kotlin/collections_maps_ops/test_map_values_view_removal_mutates_backing_map
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 2)
            val values = map.values
            __check((values.remove(2)).toString(), "true")
            __check((map.size).toString(), "2")
            __check((map.values.sum()).toString(), "3")
        }
