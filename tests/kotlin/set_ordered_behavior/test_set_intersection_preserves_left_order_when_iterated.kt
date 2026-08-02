// vybe-test: kotlin/set_ordered_behavior/test_set_intersection_preserves_left_order_when_iterated
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = linkedSetOf(1, 2, 3, 4)
            val right = linkedSetOf(4, 2)
            val inter = left.intersect(right)
            __check((inter.joinToString(",")).toString(), "2,4")
        }
