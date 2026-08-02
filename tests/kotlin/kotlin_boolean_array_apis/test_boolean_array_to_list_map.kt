// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_to_list_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun main() {
            val data = booleanArrayOf(true, true, false)
            var ones = ""
            for (i in data.indices) {
                ones = ones + data[i].toString()
                if (i + 1 < data.size) { ones = ones + "," }
            }
            println(ones)
        }

