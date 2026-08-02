// vybe-test: kotlin/java_io/test_java_io_filtered_input_stream_passthrough_count
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = java.io.BufferedInputStream(java.io.ByteArrayInputStream("aa".toByteArray()))
            val filtered = object : java.io.FilterInputStream(input) {}
            __check((filtered.read()).toString(), "97")
            __check((filtered.read()).toString(), "97")
            __check((filtered.read()).toString(), "-1")
        }
