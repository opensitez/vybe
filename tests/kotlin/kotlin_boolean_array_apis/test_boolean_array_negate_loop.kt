// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_negate_loop
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun main() {
            val data = booleanArrayOf(true, false)
            var bits = ""
            for (v in data) {
                bits = bits + if (v) "1" else "0"
            }
            println(bits)
        }

