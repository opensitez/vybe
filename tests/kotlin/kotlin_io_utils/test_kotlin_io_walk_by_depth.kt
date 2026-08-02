// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_walk_by_depth
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_walk_depth_" + System.nanoTime())
            val level1 = java.io.File(parent, "level1")
            val level2 = java.io.File(level1, "level2")
            level2.mkdirs()
            java.io.File(level2, "leaf.txt").writeText("ok")
            val names = parent.walkBottomUp().map { it.name }.toList()
            __check((names.contains("leaf.txt")).toString(), "true")
            __check((names.size).toString(), "4")
            java.io.File(level2, "leaf.txt").delete()
            level2.delete()
            level1.delete()
            parent.delete()
        }
