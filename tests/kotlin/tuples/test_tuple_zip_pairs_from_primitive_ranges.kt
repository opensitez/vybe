// vybe-test: kotlin/tuples/test_tuple_zip_pairs_from_primitive_ranges
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

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
            val zipped = (1..4).zip(10 downTo 7)
            __p((zipped.size).toString())
            __p((zipped[0]).toString())
            __p((zipped[3]).toString())
        
__check("4\n(1, 10)\n(4, 7)")
}
