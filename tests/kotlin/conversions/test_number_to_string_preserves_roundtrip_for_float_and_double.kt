// vybe-test: kotlin/conversions/test_number_to_string_preserves_roundtrip_for_float_and_double
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

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
            val a = 1.0f
            val b = 2.5f
            val d = 3.5
            val fromStringToFloat = a.toString().toFloat()
            val fromStringToDouble = d.toString().toDouble()
            __p((a.toString()).toString())
            __p((fromStringToFloat).toString())
            __p((fromStringToDouble).toString())
            __p((b.toString()).toString())
            __p((d.toString()).toString())
        
__check("1.0\n1.0\n3.5\n2.5\n3.5")
}
