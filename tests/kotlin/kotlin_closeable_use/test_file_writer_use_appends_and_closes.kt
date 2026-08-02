// vybe-test: kotlin/kotlin_closeable_use/test_file_writer_use_appends_and_closes
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val root = java.lang.System.getProperty("java.io.tmpdir")
            val file = java.io.File(root, "vybe_closeable_file_" + System.nanoTime() + ".txt")
            file.createNewFile()
            file.writeText("start")
            java.io.FileWriter(file, true).use { out ->
                out.write("-end")
            }
            val afterWrite = file.readText()
            val len = file.length().toString()
            file.delete()
            __check((afterWrite).toString(), "start-end")
            __check((len == "8").toString(), "true")
        }
