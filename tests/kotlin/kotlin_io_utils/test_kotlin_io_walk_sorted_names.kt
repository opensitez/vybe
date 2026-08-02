// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_walk_sorted_names
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_sorted_" + System.nanoTime())
            parent.mkdirs()
            java.io.File(parent, "c.txt").writeText("3")
            java.io.File(parent, "a.txt").writeText("1")
            java.io.File(parent, "b.txt").writeText("2")
            val names = parent.walk().filter { it.isFile }.map { it.name }.sorted().joinToString(",")
            __check((names).toString(), "a.txt,b.txt,c.txt")
            java.io.File(parent, "a.txt").delete()
            java.io.File(parent, "b.txt").delete()
            java.io.File(parent, "c.txt").delete()
            parent.delete()
        }
