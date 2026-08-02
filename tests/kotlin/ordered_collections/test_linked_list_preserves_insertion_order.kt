// vybe-test: kotlin/ordered_collections/test_linked_list_preserves_insertion_order
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = mutableListOf(3, 1, 2)
            __check((list.joinToString(",")).toString(), "3,1,2")
            list.add(0, 9)
            __check((list.joinToString(",")).toString(), "9,3,1,2")
        }
