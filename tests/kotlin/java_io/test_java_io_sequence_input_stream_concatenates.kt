// vybe-test: kotlin/java_io/test_java_io_sequence_input_stream_concatenates
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
            val first = java.io.ByteArrayInputStream("ab".toByteArray())
            val second = java.io.ByteArrayInputStream("cd".toByteArray())
            val seq = java.io.SequenceInputStream(first, second)
            __p((seq.read().toChar()).toString())
            __p((seq.read().toChar()).toString())
            __p((seq.read().toChar()).toString())
            __p((seq.read().toChar()).toString())
            __p((seq.read()).toString())
        
__check("a\nb\nc\nd\n-1")
}
