// vybe-test: kotlin/collections_maps_ops/test_map_filter_not_removes_matching_entries
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            val evenValues = map.filterNot { entry -> entry.value % 2 == 0 }
            __check((evenValues.keys.joinToString(",")).toString(), "a,c")
            __check((evenValues.size).toString(), "2")
        }
