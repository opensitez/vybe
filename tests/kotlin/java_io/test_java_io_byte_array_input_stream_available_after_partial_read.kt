// vybe-test: kotlin/java_io/test_java_io_byte_array_input_stream_available_after_partial_read
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = java.io.ByteArrayInputStream("1234".toByteArray())
            __check((input.read()).toString(), "49")
            __check((input.available()).toString(), "3")
        }
