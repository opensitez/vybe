// vybe-test: kotlin/ordered_collections/test_map_put_all_preserves_new_order_for_new_keys
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1)
            map.putAll(mapOf("c" to 3, "b" to 2))
            __check((map.keys.joinToString(",")).toString(), "a,c,b")
        }
