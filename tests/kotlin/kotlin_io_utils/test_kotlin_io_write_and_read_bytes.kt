// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_write_and_read_bytes
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_bytes_" + System.nanoTime() + ".bin")
            file.writeBytes(byteArrayOf(1, 2, 3, 4))
            val bytes = file.readBytes()
            __check((bytes.joinToString(",")).toString(), "1,2,3,4")
            file.delete()
        }
