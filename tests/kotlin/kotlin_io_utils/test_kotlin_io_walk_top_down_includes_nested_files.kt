// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_walk_top_down_includes_nested_files
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_walk_" + System.nanoTime())
            val nested = java.io.File(parent, "nested")
            nested.mkdirs()
            java.io.File(parent, "root.txt").writeText("r")
            java.io.File(nested, "leaf.txt").writeText("l")
            val names = parent.walkTopDown().map { it.name }.toList().sorted()
            __check((names.contains("nested")).toString(), "true")
            __check((names.contains("leaf.txt")).toString(), "true")
            __check((names.contains("root.txt")).toString(), "true")
            java.io.File(parent, "root.txt").delete()
            java.io.File(nested, "leaf.txt").delete()
            nested.delete()
            parent.delete()
        }
