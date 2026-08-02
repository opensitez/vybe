// vybe-test: kotlin/collections_sequences/test_sequence_with_generator_and_take
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun main() {
            var calls = 0
            val seq = sequence {
                var x = 0
                while (x < 5) {
                    yield(x)
                    calls += 1
                    x += 1
                }
            }
            println(seq.take(3).toList().joinToString(","))
            println(calls)
        }

