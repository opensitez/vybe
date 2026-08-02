// vybe-test: kotlin/random/test_random_next_int_range_rejects_invalid_order
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun main() {
            try {
                kotlin.random.Random(83).nextInt(10, 2)
                println("ok")
            } catch (ex: IllegalArgumentException) {
                println("error")
            }
        }

