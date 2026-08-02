// vybe-test: kotlin/java_io/test_java_io_print_writer_appendable_and_flush
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = java.io.ByteArrayOutputStream()
            val writer = java.io.PrintWriter(bytes)
            writer.append("first")
            writer.append('-')
            writer.println("second")
            writer.flush()
            __check((bytes.toString()).toString(), "first-second\n")
        }
