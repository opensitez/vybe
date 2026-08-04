// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_walk_with_depth_limit
// origin: languages/kotlin/tests/kotlin/test_kotlin_io_utils.rs

var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_depth_" + System.nanoTime())
            val d1 = java.io.File(parent, "d1")
            val d2 = java.io.File(d1, "d2")
            d2.mkdirs()
            java.io.File(d2, "f1.txt").writeText("f")
            __p((parent.walkTopDown().maxDepth(1).count()).toString())
            __p((parent.walkTopDown().maxDepth(3).count()).toString())
            java.io.File(d2, "f1.txt").delete()
            d2.delete()
            d1.delete()
            parent.delete()
        
__check("2\n4")
}
