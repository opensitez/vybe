// vybe-test: kotlin/ordered_collections/test_flatten_map_order
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(listOf(2, 1), listOf(4, 3))
            val out = values.flatten()
            __check((out.joinToString(",")).toString(), "2,1,4,3")
        }
