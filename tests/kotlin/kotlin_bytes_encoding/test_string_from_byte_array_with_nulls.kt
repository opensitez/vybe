// vybe-test: kotlin/kotlin_bytes_encoding/test_string_from_byte_array_with_nulls
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
            val bytes = byteArrayOf(78, 117, 108, 108, 0, 65)
            val value = String(bytes)
            __p((value.length).toString())
            __p((value[0]).toString())
            __p((value[4].code).toString())
        
__check("6\nN\n0")
}
