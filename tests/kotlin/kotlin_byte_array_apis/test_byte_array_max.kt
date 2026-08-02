// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_max
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun main() {
            val data = byteArrayOf(1, 5, 3)
            var max = data[0].toInt()
            for (i in 1 until data.size) {
                if (data[i].toInt() > max) { max = data[i].toInt() }
            }
            println(max)
        }

