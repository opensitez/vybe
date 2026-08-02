// vybe-test: kotlin/java_io/test_java_io_byte_array_input_stream_read_single_bytes
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = java.io.ByteArrayInputStream("xyz".toByteArray())
            __check((input.read()).toString(), "120")
            __check((input.read()).toString(), "121")
            __check((input.read()).toString(), "122")
            __check((input.read()).toString(), "-1")
        }
