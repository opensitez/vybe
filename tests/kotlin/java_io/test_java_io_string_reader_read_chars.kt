// vybe-test: kotlin/java_io/test_java_io_string_reader_read_chars
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val reader = java.io.StringReader("k1")
            val buf = CharArray(2)
            __check((reader.read(buf)).toString(), "2")
            __check((buf.joinToString(",")).toString(), "k,1")
        }
