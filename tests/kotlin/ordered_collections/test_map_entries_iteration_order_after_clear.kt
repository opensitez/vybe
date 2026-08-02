// vybe-test: kotlin/ordered_collections/test_map_entries_iteration_order_after_clear
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            map.clear()
            map["c"] = 3
            map["d"] = 4
            __check((map.entries.joinToString(",") { it.key }).toString(), "c,d")
        }
