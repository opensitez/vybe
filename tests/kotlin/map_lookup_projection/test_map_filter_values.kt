// vybe-test: kotlin/map_lookup_projection/test_map_filter_values
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
            val source = mapOf("x" to 3, "y" to 12, "z" to 18)
            val filtered = source.filterValues { it % 6 == 0 }
            __p((filtered.keys.joinToString(",")).toString())
            __p((filtered["z"]).toString())
        
__check("y,z\n18")
}
