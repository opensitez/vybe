// vybe-test: kotlin/kotlin_char_array_apis/test_char_array_upper_hint
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_array_apis.rs

fun main() {
            val a = charArrayOf('a', 'b', 'c')
            var u = ""
            for (c in a) { u = u + c.toString().uppercase() }
            println(u)
        }

