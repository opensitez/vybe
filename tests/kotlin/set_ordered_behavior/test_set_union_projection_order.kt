// vybe-test: kotlin/set_ordered_behavior/test_set_union_projection_order
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = linkedSetOf(4, 1)
            val b = linkedSetOf(2, 3)
            val all = a union b
            __check((all.joinToString(",")).toString(), "4,1,2,3")
            __check((all.size).toString(), "4")
        }
