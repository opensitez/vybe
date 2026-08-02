// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_file_permissions_approx
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_perm_" + System.nanoTime() + ".txt")
            file.writeText("perm")
            __check((file.canRead()).toString(), "true")
            __check((file.canWrite()).toString(), "true")
            file.delete()
        }
