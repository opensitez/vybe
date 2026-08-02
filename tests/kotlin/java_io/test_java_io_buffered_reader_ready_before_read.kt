// vybe-test: kotlin/java_io/test_java_io_buffered_reader_ready_before_read
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val reader = java.io.BufferedReader(java.io.StringReader("ok"))
            __check((reader.ready()).toString(), "true")
            __check((reader.read()).toString(), "111")
            __check((reader.read()).toString(), "107")
            __check((reader.read()).toString(), "-1")
        }
