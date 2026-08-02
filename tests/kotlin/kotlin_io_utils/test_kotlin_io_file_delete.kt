// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_file_delete
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_delete_" + System.nanoTime() + ".txt")
            file.writeText("gone")
            val before = file.exists()
            val deleted = file.delete()
            __check((before).toString(), "true")
            __check((deleted).toString(), "true")
            __check((file.exists()).toString(), "false")
        }
