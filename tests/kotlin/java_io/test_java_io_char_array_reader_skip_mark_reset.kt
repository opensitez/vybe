// vybe-test: kotlin/java_io/test_java_io_char_array_reader_skip_mark_reset
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val reader = java.io.CharArrayReader("abc".toCharArray())
            val one = reader.read()
            reader.mark(2)
            reader.skip(1)
            reader.reset()
            val afterReset = reader.read()
            __check((one).toString(), "97")
            __check((afterReset).toString(), "98")
        }
