// vybe-test: kotlin/arrays_ops/test_byte_array_fill_and_sum
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun main() {
            val bytes = byteArrayOf(1, 2, 3)
            bytes.fill(9.toByte(), 0, 2)
            var total = 0
            for (b in bytes) {
                total += b.toInt()
            }
            println(bytes.joinToString(","))
            println(total)
        }

