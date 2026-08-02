// vybe-test: kotlin/kotlin_short_array_apis/test_short_array_range_loop
// origin: languages/kotlin/tests/kotlin/test_kotlin_short_array_apis.rs

fun main() {
            val data = shortArrayOf(1, 2, 3)
            var out = ""
            for (i in data.indices) {
                out = out + data[i].toString()
            }
            println(out)
        }

