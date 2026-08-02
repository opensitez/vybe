// vybe-test: kotlin/set_ordered_behavior/test_linked_hash_set_keeps_insertion_and_removal_order
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = linkedSetOf(3, 1, 2, 4)
            values.remove(1)
            values.add(1)
            __check((values.joinToString(",")).toString(), "3,2,4,1")
        }
