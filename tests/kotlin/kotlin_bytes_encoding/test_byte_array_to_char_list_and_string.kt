// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_to_char_list_and_string
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

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
            val values = "AZ09".toByteArray()
            val chars = values.map { it.toInt().toChar() }.joinToString(",")
            __p((chars).toString())
            val rebuilt = String(byteArrayOf(65, 90, 48, 57))
            __p((rebuilt).toString())
        
__check("A,Z,0,9\nAZ09")
}
