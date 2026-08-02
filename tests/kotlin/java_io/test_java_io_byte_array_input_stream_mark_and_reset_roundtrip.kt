// vybe-test: kotlin/java_io/test_java_io_byte_array_input_stream_mark_and_reset_roundtrip
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = java.io.ByteArrayInputStream("abcd".toByteArray())
            __check((input.markSupported()).toString(), "true")
            __check((input.read()).toString(), "97")
            input.mark(3)
            __check((input.read()).toString(), "98")
            __check((input.read()).toString(), "99")
            input.reset()
            __check((input.read()).toString(), "98")
        }
