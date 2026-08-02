// vybe-test: kotlin/collections_sequences/test_sequence_element_at_or_else_throws_and_recovers
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun main() {
            try {
                println((1..3).asSequence().elementAt(5))
            } catch (e: IndexOutOfBoundsException) {
                println("oor")
            }
            println((1..3).asSequence().elementAtOrNull(5) ?: "none")
        }

