// vybe-test: kotlin/kotlin_sequences_generate/test_sequence_from_function_with_initial_state
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

fun main() {
            val values = sequence {
                var i = 0
                while (i < 3) {
                    yield(i)
                    i += 1
                }
            }
            println(values.map { it + 1 }.joinToString(","))
        }

