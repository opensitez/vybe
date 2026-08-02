// vybe-test: kotlin/java_io/test_java_io_buffered_input_stream_read_with_block
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = java.io.BufferedInputStream(java.io.ByteArrayInputStream("zz".toByteArray()))
            val buf = ByteArray(1)
            __check((input.read(buf)).toString(), "1")
            __check((buf[0].toInt()).toString(), "122")
        }
