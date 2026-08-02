// vybe-test: kotlin/collections_sequences/test_sequence_map_with_stateful_side_effect
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var seen = 0
            val seq = (1..5).asSequence().map { n ->
                seen += 1
                n * 10
            }
            __check(("start").toString(), "start")
            __check((seq.take(3).toList().joinToString(",")).toString(), "10,20,30")
            __check((seen).toString(), "3")
            __check((seq.toList().size).toString(), "2")
            __check((seen).toString(), "5")
        }
