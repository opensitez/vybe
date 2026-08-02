// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_temporary_file_delete
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val temp = java.io.File.createTempFile("vybe_io_ttl", ".tmp")
            temp.writeText("ttl")
            temp.deleteOnExit()
            __check((temp.exists()).toString(), "true")
            val deleted = temp.delete()
            __check((deleted).toString(), "true")
            __check((temp.exists()).toString(), "false")
        }
