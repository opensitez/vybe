// vybe-test: kotlin/java_io/test_java_io_filter_writer_passthrough
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
            val sink = java.io.StringWriter()
            val filter = object : java.io.FilterWriter(sink) {
                override fun write(i: Int) {
                    super.write(i)
                }
                override fun write(cbuf: CharArray, off: Int, len: Int) {
                    super.write(cbuf, off, len)
                }
                override fun write(str: String, off: Int, len: Int) {
                    super.write(str, off, len)
                }
                override fun flush() {
                    super.flush()
                }
                override fun close() {
                    super.close()
                }
            }
            filter.write("hello")
            filter.flush()
            __p((sink.toString()).toString())
        
__check("hello")
}
