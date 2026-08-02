// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_sum_in_long_accumulator
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun main() {
            val values = byteArrayOf(100, 101, 102)
            var total: Long = 0
            for (value in values) {
                total += value.toLong()
            }
            println(total)
            println(total > 300)
        }

