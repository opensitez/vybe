// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_walk_top_down_includes_nested_files
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
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_walk_" + System.nanoTime())
            val nested = java.io.File(parent, "nested")
            nested.mkdirs()
            java.io.File(parent, "root.txt").writeText("r")
            java.io.File(nested, "leaf.txt").writeText("l")
            val names = parent.walkTopDown().map { it.name }.toList().sorted()
            __p((names.contains("nested")).toString())
            __p((names.contains("leaf.txt")).toString())
            __p((names.contains("root.txt")).toString())
            java.io.File(parent, "root.txt").delete()
            java.io.File(nested, "leaf.txt").delete()
            nested.delete()
            parent.delete()
        
__check("true\ntrue\ntrue")
}
