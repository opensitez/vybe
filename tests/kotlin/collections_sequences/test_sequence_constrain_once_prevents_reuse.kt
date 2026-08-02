// vybe-test: kotlin/collections_sequences/test_sequence_constrain_once_prevents_reuse
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun main() {
            val constrained = sequenceOf(1, 2, 3).constrainOnce()
            println(constrained.toList().joinToString(","))
            try {
                println(constrained.toList().joinToString(","))
            } catch (e: IllegalStateException) {
                println("cannot_reuse")
            }
        }

