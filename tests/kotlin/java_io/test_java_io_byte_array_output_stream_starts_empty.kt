// vybe-test: kotlin/java_io/test_java_io_byte_array_output_stream_starts_empty
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = java.io.ByteArrayOutputStream()
            __check((out.size()).toString(), "0")
            __check((out.toString()).toString(), "")
        }
