// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_walk_with_depth_limit
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_depth_" + System.nanoTime())
            val d1 = java.io.File(parent, "d1")
            val d2 = java.io.File(d1, "d2")
            d2.mkdirs()
            java.io.File(d2, "f1.txt").writeText("f")
            __check((parent.walkTopDown().maxDepth(1).count()).toString(), "2")
            __check((parent.walkTopDown().maxDepth(3).count()).toString(), "4")
            java.io.File(d2, "f1.txt").delete()
            d2.delete()
            d1.delete()
            parent.delete()
        }
