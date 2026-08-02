// vybe-test: kotlin/collections_sequences/test_sequence_take_last_and_drop_last
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = (1..5).asSequence()
            __check((source.takeLast(3).toList().joinToString(",")).toString(), "3,4,5")
            __check((source.dropLast(1).toList().joinToString(",")).toString(), "1,2,3,4")
        }
