// vybe-test: kotlin/java_io/test_java_io_input_stream_reader_line_reading
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = java.io.ByteArrayInputStream("a\nb\n".toByteArray())
            val reader = java.io.BufferedReader(java.io.InputStreamReader(bytes))
            __check((reader.readLine()).toString(), "a")
            __check((reader.readLine()).toString(), "b")
            __check((reader.readLine() == null).toString(), "true")
        }
