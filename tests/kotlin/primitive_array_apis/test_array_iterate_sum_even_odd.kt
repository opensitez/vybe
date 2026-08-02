// vybe-test: kotlin/primitive_array_apis/test_array_iterate_sum_even_odd
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun main() {
            val values = longArrayOf(1, 2, 3, 4)
            var evens = 0
            var odds = 0
            for (v in values) {
                if (v % 2L == 0L) evens += 1 else odds += 1
            }
            println(evens)
            println(odds)
        }

