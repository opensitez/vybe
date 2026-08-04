// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_walk_sorted_names
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
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_sorted_" + System.nanoTime())
            parent.mkdirs()
            java.io.File(parent, "c.txt").writeText("3")
            java.io.File(parent, "a.txt").writeText("1")
            java.io.File(parent, "b.txt").writeText("2")
            val names = parent.walk().filter { it.isFile }.map { it.name }.sorted().joinToString(",")
            __p((names).toString())
            java.io.File(parent, "a.txt").delete()
            java.io.File(parent, "b.txt").delete()
            java.io.File(parent, "c.txt").delete()
            parent.delete()
        
__check("a.txt,b.txt,c.txt")
}
