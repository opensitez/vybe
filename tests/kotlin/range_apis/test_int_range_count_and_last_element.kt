// vybe-test: kotlin/range_apis/test_int_range_count_and_last_element
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun main() {
            val r = 2..9
            var count = 0
            var last = 0
            for (v in r) { count++
last = v }
            println(count)
            println(last)
        }

