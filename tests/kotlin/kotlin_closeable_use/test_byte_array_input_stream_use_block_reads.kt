// vybe-test: kotlin/kotlin_closeable_use/test_byte_array_input_stream_use_block_reads
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

import java.io.ByteArrayInputStream

        fun main() {
            val stream = ByteArrayInputStream("abc".toByteArray())
            val text = stream.use { s ->
                val first = s.read()
                val second = s.read()
                s.available().toString() + "," + first.toChar() + "," + second.toChar()
            }
            println(text)
            try {
                println(stream.read())
            } catch (e: Exception) {
                println("closed")
            }
        }

