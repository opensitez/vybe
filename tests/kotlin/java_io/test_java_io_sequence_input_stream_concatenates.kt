// vybe-test: kotlin/java_io/test_java_io_sequence_input_stream_concatenates
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = java.io.ByteArrayInputStream("ab".toByteArray())
            val second = java.io.ByteArrayInputStream("cd".toByteArray())
            val seq = java.io.SequenceInputStream(first, second)
            __check((seq.read().toChar()).toString(), "a")
            __check((seq.read().toChar()).toString(), "b")
            __check((seq.read().toChar()).toString(), "c")
            __check((seq.read().toChar()).toString(), "d")
            __check((seq.read()).toString(), "-1")
        }
