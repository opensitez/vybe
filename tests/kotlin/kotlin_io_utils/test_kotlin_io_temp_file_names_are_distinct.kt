// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_temp_file_names_are_distinct
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.io.File.createTempFile("vybe", "io")
            val b = java.io.File.createTempFile("vybe", "io")
            __check((a.name != b.name).toString(), "true")
            __check((a.exists()).toString(), "true")
            __check((b.exists()).toString(), "true")
            a.delete()
            b.delete()
        }
