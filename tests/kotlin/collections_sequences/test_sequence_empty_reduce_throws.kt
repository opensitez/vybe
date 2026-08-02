// vybe-test: kotlin/collections_sequences/test_sequence_empty_reduce_throws
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun main() {
            try {
                println((emptySequence<Int>()).reduce { acc, value -> acc + value })
            } catch (e: Exception) {
                println("error")
            }
        }

