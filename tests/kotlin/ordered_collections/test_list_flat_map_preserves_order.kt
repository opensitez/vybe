// vybe-test: kotlin/ordered_collections/test_list_flat_map_preserves_order
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(listOf(1, 2), listOf(3), listOf(4, 5))
            __check((values.flatMap { it }.joinToString(",")).toString(), "1,2,3,4,5")
        }
