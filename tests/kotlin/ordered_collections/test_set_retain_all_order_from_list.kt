// vybe-test: kotlin/ordered_collections/test_set_retain_all_order_from_list
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = linkedSetOf(1, 2, 3, 4)
            set.retainAll(listOf(4, 2))
            __check((set.joinToString(",")).toString(), "2,4")
        }
