// vybe-test: kotlin/random/test_random_next_int_exclusive_bound_rejects_invalid
// origin: languages/kotlin/tests/kotlin/test_random.rs

fun main() {
            try {
                kotlin.random.Random(79).nextInt(0)
                println("ok")
            } catch (ex: IllegalArgumentException) {
                println("error")
            }
        }

