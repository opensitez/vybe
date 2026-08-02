// vybe-test: kotlin/ordered_collections/test_list_set_intersection_preserves_list_order_of_first
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 2, 3, 4)
            val b = setOf(4, 2)
            val out = a.filter { b.contains(it) }
            __check((out.joinToString(",")).toString(), "2,4")
        }
