// vybe-test: kotlin/set_ordered_behavior/test_linked_set_reinsert_existing_is_noop_for_order
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = linkedSetOf(1, 2, 3)
            values.add(2)
            __check((values.joinToString(",")).toString(), "1,2,3")
        }
