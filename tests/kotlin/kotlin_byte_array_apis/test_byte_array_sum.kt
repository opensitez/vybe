// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun main() {
            val data = byteArrayOf(1, 2, 3)
            var total: Int = 0
            for (v in data) { total = total + v.toInt() }
            println(total)
        }

