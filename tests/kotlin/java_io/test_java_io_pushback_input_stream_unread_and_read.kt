// vybe-test: kotlin/java_io/test_java_io_pushback_input_stream_unread_and_read
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = java.io.ByteArrayInputStream("ab".toByteArray())
            val stream = java.io.PushbackInputStream(base)
            __check((stream.read().toChar()).toString(), "a")
            stream.unread('c'.code)
            __check((stream.read()).toString(), "99")
            __check((stream.read()).toString(), "98")
        }
