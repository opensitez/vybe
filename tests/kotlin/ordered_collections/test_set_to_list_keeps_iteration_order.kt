// vybe-test: kotlin/ordered_collections/test_set_to_list_keeps_iteration_order
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = linkedSetOf(4, 1, 3)
            __check((set.toList().joinToString(",")).toString(), "4,1,3")
            __check((set.toMutableList().sorted().joinToString(",")).toString(), "1,3,4")
        }
