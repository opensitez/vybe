// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_reader_writer_round_trip
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_rw_" + System.nanoTime() + ".txt")
            val writer = java.io.OutputStreamWriter(file.outputStream())
            writer.write("r")
            writer.write("w")
            writer.close()
            val reader = java.io.InputStreamReader(file.inputStream())
            val text = reader.readText()
            reader.close()
            __check((text).toString(), "rw")
            file.delete()
        }
