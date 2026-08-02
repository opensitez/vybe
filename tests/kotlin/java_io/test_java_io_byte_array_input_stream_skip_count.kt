// vybe-test: kotlin/java_io/test_java_io_byte_array_input_stream_skip_count
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = java.io.ByteArrayInputStream("abcdef".toByteArray())
            __check((input.skip(2)).toString(), "2")
            __check((input.read()).toString(), "99")
        }
