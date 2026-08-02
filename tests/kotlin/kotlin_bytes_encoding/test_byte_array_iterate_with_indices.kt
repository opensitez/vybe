// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_iterate_with_indices
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun main() {
            val values = byteArrayOf(4, 5, 6)
            var out = ""
            for (item in values.withIndex()) {
                out += item.index.toString() + ":" + item.value.toString() + ";"
            }
            println(out)
        }

