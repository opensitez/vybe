// vybe-test: kotlin/set_ordered_behavior/test_linked_hash_set_to_typed_array
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = linkedSetOf(7, 1, 9)
            val arr = values.toTypedArray()
            __check((arr.joinToString(",")).toString(), "7,1,9")
            val round = arr.toList().toMutableSet()
            round.add(4)
            __check((round.joinToString(",")).toString(), "7,1,9,4")
        }
