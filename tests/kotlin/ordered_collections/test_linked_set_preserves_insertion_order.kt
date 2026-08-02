// vybe-test: kotlin/ordered_collections/test_linked_set_preserves_insertion_order
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = LinkedHashSet<Int>()
            set.add(3)
            set.add(1)
            set.add(2)
            __check((set.joinToString(",")).toString(), "3,1,2")
            set.remove(1)
            set.add(1)
            __check((set.joinToString(",")).toString(), "3,2,1")
        }
