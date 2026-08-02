// vybe-test: kotlin/inline_functions/test_inline_sequence_step
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun <T> firstOrFallback(values: List<T>, fallback: T): T {
            for (item in values) {
                return item
            }
            return fallback
        }

        fun main() {
            println(firstOrFallback(listOf(8, 9), 0))
            println(firstOrFallback(emptyList(), 3))
        }

