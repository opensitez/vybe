// vybe-test: kotlin/collections_sequences/test_sequence_for_each_executes_for_all_elements_and_maintains_order
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun main() {
            var seen = ""
            (1..4).asSequence().forEach { seen += it.toString() }
            println(seen)
        }

