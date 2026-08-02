// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_is_file_and_is_directory
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_kind_" + System.nanoTime() + ".txt")
            file.writeText("kind")
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_kind_dir_" + System.nanoTime())
            dir.mkdirs()
            __check((file.isFile()).toString(), "true")
            __check((dir.isDirectory()).toString(), "true")
            file.delete()
            dir.delete()
        }
