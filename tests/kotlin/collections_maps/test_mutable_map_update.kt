// vybe-test: kotlin/collections_maps/test_mutable_map_update
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

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
            val counters = mutableMapOf("a" to 1, "b" to 2)
            counters["a"] = 4
            __p((counters["a"]).toString())
            counters.remove("b")
            __p((counters.containsKey("b")).toString())
            __p((counters.size).toString())
        
__check("4\nfalse\n1")
}
