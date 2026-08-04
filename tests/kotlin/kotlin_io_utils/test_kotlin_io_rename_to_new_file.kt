// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_rename_to_new_file
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
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"))
            val src = java.io.File(dir, "vybe_io_rename_src_" + System.nanoTime() + ".txt")
            val dst = java.io.File(dir, "vybe_io_rename_dst_" + System.nanoTime() + ".txt")
            src.writeText("rename")
            val ok = src.renameTo(dst)
            __p((ok).toString())
            __p((src.exists()).toString())
            __p((dst.readText()).toString())
            dst.delete()
        
__check("true\nfalse\nrename")
}
