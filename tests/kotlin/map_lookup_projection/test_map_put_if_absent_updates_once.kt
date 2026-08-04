// vybe-test: kotlin/map_lookup_projection/test_map_put_if_absent_updates_once
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
            val source = linkedMapOf("a" to 1)
            val existing = source.putIfAbsent("a", 99)
            val added = source.putIfAbsent("b", 2)
            __p((existing).toString())
            __p((added).toString())
            __p((source["a"]).toString())
            __p((source["b"]).toString())
        
__check("1\nnull\n1\n2")
}
