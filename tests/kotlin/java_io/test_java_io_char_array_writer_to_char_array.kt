// vybe-test: kotlin/java_io/test_java_io_char_array_writer_to_char_array
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val writer = java.io.CharArrayWriter()
            writer.write('x')
            writer.write("yz", 0, 2)
            val chars = writer.toCharArray()
            __check((chars.size).toString(), "3")
            __check((String(chars)).toString(), "xyz")
        }
