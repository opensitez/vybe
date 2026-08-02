// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_copy_to_duplicate_file
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"))
            val src = java.io.File(dir, "vybe_io_copy_src_" + System.nanoTime() + ".txt")
            val dst = java.io.File(dir, "vybe_io_copy_dst_" + System.nanoTime() + ".txt")
            src.writeText("copy")
            val copied = src.copyTo(dst, overwrite = true)
            __check((copied.readText()).toString(), "copy")
            __check((src.readText() == dst.readText()).toString(), "true")
            src.delete()
            dst.delete()
        }
