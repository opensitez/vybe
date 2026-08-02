// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_rename_to_new_file
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"))
            val src = java.io.File(dir, "vybe_io_rename_src_" + System.nanoTime() + ".txt")
            val dst = java.io.File(dir, "vybe_io_rename_dst_" + System.nanoTime() + ".txt")
            src.writeText("rename")
            val ok = src.renameTo(dst)
            __check((ok).toString(), "true")
            __check((src.exists()).toString(), "false")
            __check((dst.readText()).toString(), "rename")
            dst.delete()
        }
