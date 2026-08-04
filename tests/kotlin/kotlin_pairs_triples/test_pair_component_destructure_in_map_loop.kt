// vybe-test: kotlin/kotlin_pairs_triples/test_pair_component_destructure_in_map_loop
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

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
            val map = mapOf("x" to 10, "y" to 20)
            var total = 0
            for ((k, v) in map) {
                total += if (k == "x") v else 0
            }
            __p((total).toString())
        
__check("10")
}
