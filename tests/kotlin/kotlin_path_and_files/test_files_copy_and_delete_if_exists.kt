// vybe-test: kotlin/kotlin_path_and_files/test_files_copy_and_delete_if_exists
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Files
        import java.nio.file.StandardCopyOption

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
            val src = Files.createTempFile("vybe-copy-src", ".txt")
            Files.writeString(src, "alpha")
            val dst = Files.createTempFile("vybe-copy-dst", ".txt")
            Files.copy(src, dst, StandardCopyOption.REPLACE_EXISTING)
            val before = Files.readString(dst)
            val removed = Files.deleteIfExists(src)
            __p((before).toString())
            __p((removed).toString())
            __p((Files.exists(src)).toString())
            Files.delete(dst)
        
__check("alpha\ntrue\nfalse")
}
