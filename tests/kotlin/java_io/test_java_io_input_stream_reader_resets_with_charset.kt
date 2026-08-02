// vybe-test: kotlin/java_io/test_java_io_input_stream_reader_resets_with_charset
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = java.io.ByteArrayInputStream("hi".toByteArray("UTF-8"))
            val reader = java.io.InputStreamReader(bytes, java.nio.charset.StandardCharsets.UTF_8)
            __check((reader.ready()).toString(), "true")
            __check((reader.read()).toString(), "104")
            __check((reader.read()).toString(), "105")
            __check((reader.read()).toString(), "-1")
        }
