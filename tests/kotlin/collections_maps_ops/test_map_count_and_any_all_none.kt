// vybe-test: kotlin/collections_maps_ops/test_map_count_and_any_all_none
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            __check((map.count { it.value > 1 }).toString(), "2")
            __check((map.any { it.key == "b" }).toString(), "true")
            __check((map.all { it.value > 0 }).toString(), "true")
            __check((map.none { it.key == "z" }).toString(), "true")
        }
