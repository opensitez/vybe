// vybe-test: kotlin/inline_functions/test_inline_chaining_calls
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun firstNonEmpty(values: List<String>): String {
            for (value in values) {
                if (value.isNotEmpty()) return value
            }
            return ""
        }

        fun main() {
            println(firstNonEmpty(listOf("", "kotlin", "x")))
        }

