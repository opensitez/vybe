// vybe-test: kotlin/collection_projection_views/test_intersect_retains_order_in_list_result
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 3, 5, 7, 3)
            val b = listOf(3, 3, 7)
            val out1 = a.intersect(b.toSet())
            __check((out1.joinToString(",")).toString(), "3,7")
        }
