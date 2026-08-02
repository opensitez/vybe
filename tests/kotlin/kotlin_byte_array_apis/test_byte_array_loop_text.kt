// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_loop_text
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun main() {
            val data = byteArrayOf(4, 5)
            var s = ""
            for (v in data) { s = s + v.toString() }
            println(s)
        }

