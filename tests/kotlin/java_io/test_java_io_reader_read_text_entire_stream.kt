// vybe-test: kotlin/java_io/test_java_io_reader_read_text_entire_stream
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

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
            val text = "kotlin stream"
            val reader = java.io.StringReader(text)
            val writer = java.io.StringWriter()
            val buf = CharArray(4)
            while (true) {
                val count = reader.read(buf)
                if (count < 0) break
                writer.write(buf, 0, count)
            }
            __p((writer.toString()).toString())
        
__check("kotlin stream")
}
