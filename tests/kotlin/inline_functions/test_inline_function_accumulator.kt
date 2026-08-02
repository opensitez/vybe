// vybe-test: kotlin/inline_functions/test_inline_function_accumulator
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun <T> fold(start: T, values: List<T>, op: (T, T) -> T): T {
            var out = start
            for (value in values) {
                out = op(out, value)
            }
            return out
        }

        fun main() {
            println(fold(0, listOf(1, 2, 3), { a, b -> a + b }))
        }

