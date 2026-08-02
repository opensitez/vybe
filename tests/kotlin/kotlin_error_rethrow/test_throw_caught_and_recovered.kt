// vybe-test: kotlin/kotlin_error_rethrow/test_throw_caught_and_recovered
// origin: languages/kotlin/tests/kotlin/test_kotlin_error_rethrow.rs

fun main() {
            fun mustBePositive(value: Int): Int {
                if (value <= 0) throw Exception("bad")
                return value
            }

            try {
                println(mustBePositive(2))
            } catch (e: Exception) {
                println("error")
            } finally {
                println("done")
            }
        }

