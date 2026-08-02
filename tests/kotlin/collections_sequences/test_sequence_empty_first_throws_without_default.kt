// vybe-test: kotlin/collections_sequences/test_sequence_empty_first_throws_without_default
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun main() {
            try {
                println(emptySequence<Int>().first())
            } catch (e: NoSuchElementException) {
                println("empty")
            }
        }

