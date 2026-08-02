// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_last_modified_is_time_like
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_mtime_" + System.nanoTime() + ".txt")
            file.writeText("time")
            val mtime = file.lastModified()
            __check((mtime > 0).toString(), "true")
            file.delete()
        }
