// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_append_text_extends_existing_content
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_append_" + System.nanoTime() + ".txt")
            file.writeText("a")
            file.appendText("b")
            file.appendText("c")
            __check((file.readText()).toString(), "abc")
            file.delete()
        }
