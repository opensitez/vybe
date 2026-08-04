// vybe-test: kotlin/tuples/test_tuple_pair_zip_default_no_transform_is_pair
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
            val left = listOf(1, 2, 3)
            val right = listOf("a", "b", "c")
            val zipped = left.zip(right)
            __p((zipped.size).toString())
            __p((zipped[0]).toString())
            __p((zipped[1].first + zipped[1].second.length).toString())
        
__check("3\n(1, a)\n3")
}
