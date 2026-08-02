// vybe-test: kotlin/primitive_array_apis/test_char_array_indices_iteration
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun main() {
            val values = charArrayOf('a', 'b', 'c')
            var outText = ""
            for (i in values.indices) {
                outText += i.toString() + values[i]
            }
            println(outText)
        }

