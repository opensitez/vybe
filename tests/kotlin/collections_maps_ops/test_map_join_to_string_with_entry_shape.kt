// vybe-test: kotlin/collections_maps_ops/test_map_join_to_string_with_entry_shape
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            val formatted = map.entries.joinToString("|") { "${it.key}:${it.value}" }
            __check((formatted).toString(), "a:1|b:2")
        }
