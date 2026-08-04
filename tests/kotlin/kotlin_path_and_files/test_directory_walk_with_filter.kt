// vybe-test: kotlin/kotlin_path_and_files/test_directory_walk_with_filter
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Files
        import java.nio.file.Paths

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
            val base = Paths.get(java.lang.System.getProperty("java.io.tmpdir"), "vybe_walk_" + System.nanoTime().toString())
            val a = Files.createDirectories(base.resolve("a"))
            val b = base.resolve("a.txt")
            val c = base.resolve("b.log")
            Files.writeString(b, "one")
            Files.writeString(c, "two")
            val count = Files.list(base).filter { p -> Files.isRegularFile(p) }.count().toInt()
            __p((count).toString())
            val hasTxt = Files.newDirectoryStream(base, "*.txt").use {
                it.asSequence().count()
            }
            __p((hasTxt).toString())
            Files.delete(b)
            Files.delete(c)
            Files.delete(a)
            Files.delete(base)
        
__check("2\n1")
}
