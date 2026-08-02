// vybe-test: kotlin/java_io/test_java_io_print_stream_no_auto_flush_without_newline
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val printer = java.io.PrintStream(sink)
            printer.print("no_flush")
            printer.flush()
            __check((sink.toString()).toString(), "no_flush")
        }
