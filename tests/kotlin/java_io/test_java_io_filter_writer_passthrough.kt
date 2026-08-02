// vybe-test: kotlin/java_io/test_java_io_filter_writer_passthrough
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sink = java.io.StringWriter()
            val filter = object : java.io.FilterWriter(sink) {
                override fun write(i: Int) {
                    super.write(i)
                }
                override fun write(cbuf: CharArray, off: Int, len: Int) {
                    super.write(cbuf, off, len)
                }
                override fun write(str: String, off: Int, len: Int) {
                    super.write(str, off, len)
                }
                override fun flush() {
                    super.flush()
                }
                override fun close() {
                    super.close()
                }
            }
            filter.write("hello")
            filter.flush()
            __check((sink.toString()).toString(), "hello")
        }
