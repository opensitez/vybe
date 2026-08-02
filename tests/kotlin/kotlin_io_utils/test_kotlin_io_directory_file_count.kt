// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_directory_file_count
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_count_" + System.nanoTime())
            parent.mkdirs()
            java.io.File(parent, "a").mkdir()
            java.io.File(parent, "b.txt").writeText("b")
            java.io.File(parent, "c.txt").writeText("c")
            val files = parent.listFiles { f -> f.isFile }
            __check((files.size).toString(), "2")
            val dirs = parent.listFiles { f -> f.isDirectory }
            __check((dirs.size).toString(), "1")
            java.io.File(parent, "b.txt").delete()
            java.io.File(parent, "c.txt").delete()
            java.io.File(parent, "a").delete()
            parent.delete()
        }
