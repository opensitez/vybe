// vybe-test: kotlin/collections_sequences/test_sequence_generate_sequence_side_effects_are_short_circuited_by_take
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var calls = 0
            val seq = sequence {
                yield(1)
                calls += 1
                yield(2)
                calls += 1
                yield(3)
                calls += 1
                yield(4)
            }
            __check((seq.take(3).toList().joinToString(",")).toString(), "1,2,3")
            __check((calls).toString(), "2")
        }
