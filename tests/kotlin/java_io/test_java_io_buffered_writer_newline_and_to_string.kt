// vybe-test: kotlin/java_io/test_java_io_buffered_writer_newline_and_to_string
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val stringWriter = java.io.StringWriter()
            val writer = java.io.BufferedWriter(stringWriter)
            writer.write("line1")
            writer.newLine()
            writer.write("line2")
            writer.flush()
            __check((stringWriter.toString()).toString(), "line1\nline2")
        }
