// vybe-test: kotlin/java_io/test_java_io_writer_error_state_is_false_after_success
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val writer = java.io.PrintWriter(sink)
            writer.print("x")
            writer.flush()
            __check((writer.checkError()).toString(), "false")
        }
