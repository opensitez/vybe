// vybe-test: kotlin/java_io/test_java_io_output_stream_writer_flush_no_op_if_closed
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val writer = java.io.OutputStreamWriter(sink)
            writer.write("test")
            writer.flush()
            writer.close()
            __check((sink.toString()).toString(), "test")
        }
