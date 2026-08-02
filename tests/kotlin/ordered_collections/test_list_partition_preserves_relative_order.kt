// vybe-test: kotlin/ordered_collections/test_list_partition_preserves_relative_order
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, 2, 3, 4, 5)
            val (a, b) = list.partition { it % 2 == 0 }
            __check((a.joinToString(",")).toString(), "2,4")
            __check((b.joinToString(",")).toString(), "1,3,5")
        }
