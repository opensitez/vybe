// vybe-test: kotlin/numeric_types/test_byte_and_short_roundtrip_via_int
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

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
            val b: Byte = 127
            val s: Short = 32767
            __p((b.toInt() + 1).toString())
            __p((s.toInt() + 1).toString())
            __p((b.toLong() - 7).toString())
            __p((s.toLong() - 7).toString())
        
__check("128\n32768\n120\n32760")
}
