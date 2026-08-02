// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_empty_file_reads_empty_lines
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_empty_" + System.nanoTime() + ".txt")
            file.writeText("")
            val bytes = file.readBytes()
            val lines = file.readLines()
            __check((bytes.size).toString(), "0")
            __check((lines.size).toString(), "0")
            file.delete()
        }
