// vybe-test: kotlin/kotlin_char_array_apis/test_char_array_reverse
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_array_apis.rs

fun main() {
            val data = charArrayOf('a', 'b', 'c')
            var out = ""
            for (i in data.indices.reversed()) {
                out = out + data[i].toString()
            }
            println(out)
        }

