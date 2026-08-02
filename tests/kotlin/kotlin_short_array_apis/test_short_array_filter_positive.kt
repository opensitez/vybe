// vybe-test: kotlin/kotlin_short_array_apis/test_short_array_filter_positive
// origin: languages/kotlin/tests/kotlin/test_kotlin_short_array_apis.rs

fun main() {
            val a = shortArrayOf(-1, 0, 2)
            var count = 0
            for (x in a) { if (x.toInt() > 0) { count = count + 1 } }
            println(count)
        }

