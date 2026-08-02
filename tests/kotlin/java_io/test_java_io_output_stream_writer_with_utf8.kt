// vybe-test: kotlin/java_io/test_java_io_output_stream_writer_with_utf8
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = java.io.ByteArrayOutputStream()
            val writer = java.io.OutputStreamWriter(bytes, java.nio.charset.StandardCharsets.UTF_8)
            writer.write("ß")
            writer.flush()
            __check((bytes.toString("UTF-8")).toString(), "ß")
        }
