// vybe-test: kotlin/kotlin_path_and_files/test_copy_to_multiple_targets_preserves_contents
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
            val src = Files.createTempFile("vybe-copy-a", ".txt")
            Files.writeString(src, "data")
            val d1 = Files.createTempFile("vybe-copy-b", ".txt")
            val d2 = Files.createTempFile("vybe-copy-c", ".txt")
            Files.copy(src, d1, StandardCopyOption.REPLACE_EXISTING)
            Files.copy(src, d2, StandardCopyOption.REPLACE_EXISTING)
            __p((Files.readString(d1)).toString())
            __p((Files.readString(d2)).toString())
            Files.delete(src)
            Files.delete(d1)
            Files.delete(d2)
        
__check("data\ndata")
}
