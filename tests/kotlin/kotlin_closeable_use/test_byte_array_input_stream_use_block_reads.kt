// vybe-test: kotlin/kotlin_closeable_use/test_byte_array_input_stream_use_block_reads
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

import java.io.ByteArrayInputStream

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
            val stream = ByteArrayInputStream("abc".toByteArray())
            val text = stream.use { s ->
                val first = s.read()
                val second = s.read()
                s.available().toString() + "," + first.toChar() + "," + second.toChar()
            }
            __p((text).toString())
            try {
                __p((stream.read()).toString())
            } catch (e: Exception) {
                __p(("closed").toString())
            }
        
__check("1,b,c\nclosed")
}
