// vybe-test: kotlin/kotlin_io_utils/test_kotlin_io_file_name_with_path_methods
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
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_name_" + System.nanoTime() + ".txt")
            file.writeText("x")
            __p((file.path.contains(file.name)).toString())
            __p((file.absoluteFile.name == file.name).toString())
            __p((file.toPath().fileName.toString().endsWith(".txt")).toString())
            file.delete()
        
__check("true\ntrue\ntrue")
}
