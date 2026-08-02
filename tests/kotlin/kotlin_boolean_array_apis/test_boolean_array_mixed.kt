// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_mixed
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun main() {
            val data = booleanArrayOf(true, false, false, true)
            var i = 0
            var out = ""
            while (i < data.size) {
                out = out + if (data[i]) "T" else "F"
                i = i + 1
            }
            println(out)
        }

