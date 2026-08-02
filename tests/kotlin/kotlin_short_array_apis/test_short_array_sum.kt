// vybe-test: kotlin/kotlin_short_array_apis/test_short_array_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_short_array_apis.rs

fun main() {
            val data = shortArrayOf(4, 5, 6)
            var total: Int = 0
            for (x in data) { total = total + x.toInt() }
            println(total)
        }

