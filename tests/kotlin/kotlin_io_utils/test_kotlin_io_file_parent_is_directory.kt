// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_file_parent_is_directory
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_parent_" + System.nanoTime() + ".txt")
            file.writeText("x")
            val parent = file.parentFile
            __check((parent.isDirectory()).toString(), "true")
            __check((file.toPath().parent != null).toString(), "true")
            file.delete()
        }
