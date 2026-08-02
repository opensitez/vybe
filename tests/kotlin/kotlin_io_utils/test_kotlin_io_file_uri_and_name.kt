// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_file_uri_and_name
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_uri_" + System.nanoTime() + ".txt")
            file.writeText("uri")
            val uri = file.toURI()
            __check((uri.toString().endsWith(".txt")).toString(), "true")
            __check((file.name.startsWith("vybe_io_uri_")).toString(), "true")
            file.delete()
        }
