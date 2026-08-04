// vybe-test: kotlin/map_lookup_projection/test_map_to_pair_lists_round_trip
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

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
            val source = mapOf("x" to 9, "y" to 8)
            val pairs = source.toList()
            val rebuilt = pairs.toMap()
            __p((pairs.joinToString("|") { it.toString() }).toString())
            __p((rebuilt.size).toString())
            __p((rebuilt["y"]).toString())
        
__check("(x, 9)|(y, 8)\n2\n8")
}
