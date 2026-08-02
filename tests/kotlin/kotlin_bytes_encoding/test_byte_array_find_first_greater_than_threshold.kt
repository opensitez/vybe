// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_find_first_greater_than_threshold
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun main() {
            val values = byteArrayOf(4, 6, 8, 9)
            var found = -1
            for (value in values) {
                if (value > 7) {
                    found = value.toInt()
                    break
                }
            }
            println(found)
        }

