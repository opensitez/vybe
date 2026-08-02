// vybe-test: kotlin/java_io/test_java_io_byte_array_output_stream_to_byte_array_len
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val stream = java.io.ByteArrayOutputStream()
            stream.write("abc".toByteArray())
            val data = stream.toByteArray()
            __check((data.size).toString(), "3")
            __check((data[1]).toString(), "98")
        }
