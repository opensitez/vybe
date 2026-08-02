// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_create_nested_directory_and_list_children
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_dir_" + System.nanoTime())
            parent.mkdirs()
            val childA = java.io.File(parent, "a.txt")
            val childB = java.io.File(parent, "b.txt")
            childA.writeText("1")
            childB.writeText("2")
            val names = parent.listFiles()
            val joined = names.map { it.name }.sorted().joinToString(",")
            __check((joined).toString(), "a.txt,b.txt")
            __check((parent.delete()).toString(), "false")
            childA.delete()
            childB.delete()
            parent.delete()
        }
