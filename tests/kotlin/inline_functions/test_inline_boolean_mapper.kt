// vybe-test: kotlin/inline_functions/test_inline_boolean_mapper
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun <T> allMatch(values: List<T>, check: (T) -> Boolean): Boolean {
            for (value in values) {
                if (!check(value)) return false
            }
            return true
        }

        fun main() {
            println(allMatch(listOf(2, 4, 6)) { it % 2 == 0 })
            println(allMatch(listOf(2, 3, 6)) { it % 2 == 0 })
        }

