// vybe-test: kotlin/arrays_ops/test_byte_array_join_and_sum
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun main() {
            val bytes = byteArrayOf(5, -1, 3)
            val shifted = bytes.map { it.toInt() + 1 }.toByteArray()
            var total = 0
            for (value in shifted) {
                total += value.toInt()
            }
            println(shifted.joinToString(","))
            println(total)
        }

