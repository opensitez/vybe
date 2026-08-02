// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_nested_loop
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun main() {
            val a = booleanArrayOf(true, false)
            var out = ""
            for (x in a) {
                for (y in a) {
                    out = out + if (x && y) "1" else "0"
                }
            }
            println(out)
        }

