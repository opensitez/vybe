// vybe-test: kotlin/java_io/test_java_io_print_stream_auto_flush
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val printer = java.io.PrintStream(sink, true)
            printer.println("one")
            __check((sink.toString()).toString(), "one\n")
        }
