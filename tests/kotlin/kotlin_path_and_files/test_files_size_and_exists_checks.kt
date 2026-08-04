// vybe-test: kotlin/kotlin_path_and_files/test_files_size_and_exists_checks
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Files

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
            val path = Files.createTempFile("vybe-size", ".txt")
            Files.writeString(path, "kotlin")
            __p((Files.exists(path)).toString())
            __p((Files.isRegularFile(path)).toString())
            __p((Files.size(path)).toString())
            Files.delete(path)
            __p((Files.exists(path)).toString())
        
__check("true\ntrue\n6\nfalse")
}
