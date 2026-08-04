// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_builder
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

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
            val u = UIntArray(4) { it.toUInt() + 1u }
            val b = UByteArray(3) { (it + 10).toUByte() }
            val s = UShortArray(2) { ((it * 2 + 1).toUShort()) }
            val l = ULongArray(2) { (it.toULong() + 1uL) * 100uL }
            __p((u.joinToString(",")).toString())
            __p((b.joinToString(",")).toString())
            __p((s.joinToString(",")).toString())
            __p((l.joinToString(",")).toString())
        
__check("1,2,3,4\n10,11,12\n1,3\n100,200")
}
