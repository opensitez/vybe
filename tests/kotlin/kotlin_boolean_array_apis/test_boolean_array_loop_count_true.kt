// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_loop_count_true
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun main() {
            val data = booleanArrayOf(true, false, true, true)
            var count = 0
            for (v in data) {
                if (v) { count = count + 1 }
            }
            println(count)
        }

