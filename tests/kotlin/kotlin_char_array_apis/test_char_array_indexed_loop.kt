// vybe-test: kotlin/kotlin_char_array_apis/test_char_array_indexed_loop
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_array_apis.rs

fun main() {
            val a = charArrayOf('a', 'b', 'c')
            var out = ""
            var i = 0
            while (i < a.size) {
                out = out + a[i].toString()
                i = i + 1
            }
            println(out)
        }

