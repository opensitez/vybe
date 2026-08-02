// vybe-test: kotlin/preconditions/test_require_orchestrates_guard_clause
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun clamp(value: Int): Int {
            require(value in 1..10)
            return value
        }

        fun main() {
            println(clamp(7))
            try {
                println(clamp(42))
            } catch (e: IllegalArgumentException) {
                println("bad")
            }
        }

