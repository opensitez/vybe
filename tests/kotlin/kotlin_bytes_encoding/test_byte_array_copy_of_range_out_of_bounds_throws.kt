// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_copy_of_range_out_of_bounds_throws
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun main() {
            val values = byteArrayOf(1, 2, 3)
            try {
                values.copyOfRange(4, 6)
                println("ok")
            } catch (e: Exception) {
                println(e::class.simpleName)
            }
        }

