// vybe-test: kotlin/kotlin_path_and_files/test_files_temp_write_read_delete_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_path_and_files.rs

import java.nio.file.Files
        import java.nio.file.Path

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
            val tmp = Files.createTempFile("vybe", ".txt")
            Files.writeString(tmp, "hello")
            val text = Files.readString(tmp)
            __p((text).toString())
            val moved = Files.move(tmp, tmp.resolveSibling("vybe_moved_" + tmp.fileName.toString()), java.nio.file.StandardCopyOption.REPLACE_EXISTING)
            __p((Files.exists(tmp)).toString())
            __p((Files.exists(moved)).toString())
            __p((Files.readString(moved)).toString())
            Files.delete(moved)
        
__check("hello\nfalse\ntrue\nhello")
}
