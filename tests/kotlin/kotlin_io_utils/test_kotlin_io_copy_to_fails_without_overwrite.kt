// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_copy_to_fails_without_overwrite
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun main() {
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"))
            val src = java.io.File(dir, "vybe_io_cfail_src_" + System.nanoTime() + ".txt")
            val dst = java.io.File(dir, "vybe_io_cfail_dst_" + System.nanoTime() + ".txt")
            src.writeText("src")
            dst.writeText("dst")
            try {
                src.copyTo(dst, overwrite = false)
                println("no_error")
            } catch (e: Exception) {
                println(e::class.simpleName)
            }
            src.delete()
            dst.delete()
        }

