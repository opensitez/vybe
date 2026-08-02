// vybe-test: kotlin/java_io/test_java_io_byte_array_output_stream_reset
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val stream = java.io.ByteArrayOutputStream()
            stream.write("before".toByteArray())
            stream.reset()
            stream.write("after".toByteArray())
            __check((stream.toString()).toString(), "after")
        }
