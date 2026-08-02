// vybe-test: kotlin/kotlin_char_array_apis/test_char_array_contains
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_array_apis.rs

fun main() {
            val data = charArrayOf('a', 'b', 'c')
            var hit = false
            for (c in data) { if (c == 'b') { hit = true } }
            println(hit.toString())
        }

