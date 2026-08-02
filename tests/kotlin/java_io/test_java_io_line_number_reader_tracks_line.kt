// vybe-test: kotlin/java_io/test_java_io_line_number_reader_tracks_line
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val reader = java.io.LineNumberReader(java.io.StringReader("x\n y\n"))
            __check((reader.readLine()).toString(), "x")
            __check((reader.getLineNumber()).toString(), "1")
            __check((reader.readLine()).toString(), " y")
            __check((reader.getLineNumber()).toString(), "2")
        }
