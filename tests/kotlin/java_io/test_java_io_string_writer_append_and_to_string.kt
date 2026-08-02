// vybe-test: kotlin/java_io/test_java_io_string_writer_append_and_to_string
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val writer = java.io.StringWriter()
            writer.write("a")
            writer.append('b')
            writer.append("c")
            __check((writer.toString()).toString(), "abc")
        }
