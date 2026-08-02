// vybe-test: kotlin/collections_sequences/test_sequence_reused_iterable_runs_again_when_materialized_twice
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
                calls += 1
                yield(1)
                calls += 1
                yield(2)
            }
            __check((seq.toList().joinToString(",")).toString(), "1,2")
            __check((seq.toList().joinToString(",")).toString(), "1,2")
            __check((calls).toString(), "4")
        }
