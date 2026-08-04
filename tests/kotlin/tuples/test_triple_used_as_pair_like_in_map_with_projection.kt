// vybe-test: kotlin/tuples/test_triple_used_as_pair_like_in_map_with_projection
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
            val points = listOf(
                Triple("a", 1, 10),
                Triple("b", 2, 20)
            ).associateBy { it.first }
            val first = "a"
            val values = points[first]!!
            __p((first).toString())
            __p((values.second).toString())
            __p((values.third).toString())
        
__check("a\n1\n10")
}
