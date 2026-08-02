// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_file_name_with_path_methods
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_name_" + System.nanoTime() + ".txt")
            file.writeText("x")
            __check((file.path.contains(file.name)).toString(), "true")
            __check((file.absoluteFile.name).toString(), "true")
            __check((file.toPath().fileName.toString().endsWith(".txt")).toString(), "true")
            file.delete()
        }
