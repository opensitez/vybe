// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_absolute_and_parent_paths
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_abs_" + System.nanoTime() + ".txt")
            file.writeText("x")
            val absolute = file.absolutePath
            val parent = file.parent
            __check((absolute.startsWith(java.lang.System.getProperty("java.io.tmpdir"))).toString(), "true")
            __check((parent != null).toString(), "true")
            file.delete()
        }
