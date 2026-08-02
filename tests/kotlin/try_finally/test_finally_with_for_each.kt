// vybe-test: kotlin/try_finally/test_finally_with_for_each
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun main() {
        val data = intArrayOf(1,2)
        var sum = 0
        data.forEach { v ->
            try {
                sum += v
            } finally {
                sum += 1
            }
        }
        println(sum)
    }

