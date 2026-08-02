// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_use_lines_counts_lines
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_use_lines_" + System.nanoTime() + ".txt")
            file.writeText("1\n2\n3")
            val count = file.useLines { lines -> lines.count() }
            __check((count).toString(), "3")
            file.delete()
        }
