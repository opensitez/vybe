// vybe-test: kotlin/kotlin_char_array_apis/test_char_array_concat
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_array_apis.rs

fun main() {
            val data = charArrayOf('x', 'y')
            var s = ""
            for (c in data) { s = s + c.toString() }
            println(s)
        }

