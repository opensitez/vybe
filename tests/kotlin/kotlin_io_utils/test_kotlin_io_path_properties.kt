// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_path_properties
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_props_" + System.nanoTime() + ".dat")
            file.writeText("p")
            __check((file.name.endsWith(".dat")).toString(), "true")
            __check((file.extension).toString(), "dat")
            __check((file.nameWithoutExtension.contains("vybe_io_props_")).toString(), "true")
            file.delete()
        }
