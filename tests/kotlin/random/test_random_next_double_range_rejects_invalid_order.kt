// vybe-test: kotlin/random/test_random_next_double_range_rejects_invalid_order
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun main() {
            try {
                kotlin.random.Random(97).nextDouble(2.0, 1.0)
                println("ok")
            } catch (ex: IllegalArgumentException) {
                println("error")
            }
        }

