// vybe-test: kotlin/collections_maps_ops/test_map_entries_find_and_first
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("x" to 10, "y" to 11, "z" to 12)
            val found = map.entries.find { it.value == 11 }
            __check((found?.key ?: "none").toString(), "y")
            val first = map.entries.first()
            __check((first.key + ":" + first.value).toString(), "x:10")
        }
