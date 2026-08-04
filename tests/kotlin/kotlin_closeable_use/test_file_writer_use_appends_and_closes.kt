// vybe-test: kotlin/kotlin_closeable_use/test_file_writer_use_appends_and_closes
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

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
            val root = java.lang.System.getProperty("java.io.tmpdir")
            val file = java.io.File(root, "vybe_closeable_file_" + System.nanoTime() + ".txt")
            file.createNewFile()
            file.writeText("start")
            java.io.FileWriter(file, true).use { out ->
                out.write("-end")
            }
            val afterWrite = file.readText()
            val len = file.length().toString()
            file.delete()
            __p((afterWrite).toString())
            __p((len == "8").toString())
        
__check("start-end\ntrue")
}
