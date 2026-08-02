// vybe-test: kotlin/kotlin_char_array_apis/test_char_array_find_first
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_array_apis.rs

fun main() {
            val a = charArrayOf('x', 'y', 'z')
            var pos = -1
            for (i in a.indices) {
                if (a[i] == 'y') { pos = i }
            }
            println(pos)
        }

