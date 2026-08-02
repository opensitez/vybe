// vybe-test: kotlin/java_io/test_java_io_reader_read_text_entire_stream
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun main() {
            val text = "kotlin stream"
            val reader = java.io.StringReader(text)
            val writer = java.io.StringWriter()
            val buf = CharArray(4)
            while (true) {
                val count = reader.read(buf)
                if (count < 0) break
                writer.write(buf, 0, count)
            }
            println(writer.toString())
        }

