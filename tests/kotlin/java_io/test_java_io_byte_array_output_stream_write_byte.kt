// vybe-test: kotlin/java_io/test_java_io_byte_array_output_stream_write_byte
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val stream = java.io.ByteArrayOutputStream()
            stream.write(65)
            __check((stream.toString()).toString(), "A")
            __check((stream.size()).toString(), "1")
        }
