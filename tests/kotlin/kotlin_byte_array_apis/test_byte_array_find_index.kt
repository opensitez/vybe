// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_find_index
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun main() {
            val a = byteArrayOf(6, 7, 8)
            var idx = -1
            for (i in a.indices) {
                if (a[i] == 7.toByte()) { idx = i }
            }
            println(idx)
        }

