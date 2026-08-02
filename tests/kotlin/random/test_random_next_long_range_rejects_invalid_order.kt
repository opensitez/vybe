// vybe-test: kotlin/random/test_random_next_long_range_rejects_invalid_order
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun main() {
            try {
                kotlin.random.Random(89).nextLong(3L, -3L)
                println("ok")
            } catch (ex: IllegalArgumentException) {
                println("error")
            }
        }

