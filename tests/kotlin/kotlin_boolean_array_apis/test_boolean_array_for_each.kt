// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_for_each
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun main() {
            val data = booleanArrayOf(false, true)
            var total = ""
            data.forEach { item ->
                total = total + item.toString()
            }
            println(total)
        }

