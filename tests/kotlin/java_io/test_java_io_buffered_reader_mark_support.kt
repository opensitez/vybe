// vybe-test: kotlin/java_io/test_java_io_buffered_reader_mark_support
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val reader = java.io.BufferedReader(java.io.StringReader("12345"))
            __check((reader.markSupported()).toString(), "true")
        }
