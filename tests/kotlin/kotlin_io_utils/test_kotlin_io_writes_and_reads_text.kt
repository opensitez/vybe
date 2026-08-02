// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_writes_and_reads_text
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_text_" + System.nanoTime() + "_a.txt")
            file.writeText("hello")
            __check((file.exists()).toString(), "true")
            __check((file.readText()).toString(), "hello")
            file.delete()
        }
