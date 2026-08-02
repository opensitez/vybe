// vybe-test: kotlin/java_io/test_java_io_print_writer_writes_text
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = java.io.ByteArrayOutputStream()
            val printer = java.io.PrintWriter(bytes)
            printer.println("kotlin")
            printer.print("rocks")
            printer.flush()
            __check((bytes.toString()).toString(), "kotlin\nrocks")
        }
