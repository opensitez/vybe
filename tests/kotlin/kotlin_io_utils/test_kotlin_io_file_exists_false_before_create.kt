// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_file_exists_false_before_create
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_missing_" + System.nanoTime() + ".txt")
            __check((file.exists()).toString(), "false")
            file.writeText("now")
            __check((file.exists()).toString(), "true")
            file.delete()
        }
