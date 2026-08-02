// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_loop_sum_accumulator
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun main() {
            val values = byteArrayOf(2, 4, 6)
            var total = 0
            for (value in values) {
                total += value.toInt()
            }
            println(total)
        }

