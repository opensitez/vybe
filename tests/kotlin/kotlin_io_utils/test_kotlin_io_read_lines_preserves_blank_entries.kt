// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_read_lines_preserves_blank_entries
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_lines_" + System.nanoTime() + ".txt")
            file.writeText("a\n\n b\n")
            val lines = file.readLines()
            __check((lines.size).toString(), "3")
            __check((lines.joinToString("|")).toString(), "a|| b")
            file.delete()
        }
