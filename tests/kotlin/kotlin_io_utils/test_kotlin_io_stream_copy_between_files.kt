// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_stream_copy_between_files
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"))
            val src = java.io.File(dir, "vybe_io_stream_src_" + System.nanoTime() + ".txt")
            val dst = java.io.File(dir, "vybe_io_stream_dst_" + System.nanoTime() + ".txt")
            src.writeText("stream")
            src.inputStream().use { input ->
                dst.outputStream().use { output ->
                    input.copyTo(output)
                }
            }
            __check((dst.readText()).toString(), "stream")
            src.delete()
            dst.delete()
        }
