// vybe-test: kotlin/collections_maps_ops/test_map_merge_without_overwrite_keeps_existing
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

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
            val base = linkedMapOf("a" to 1, "b" to 2)
            val extras = mapOf("b" to 20, "c" to 3)
            val merged = extras + base
            __p((merged["a"]).toString())
            __p((merged["b"]).toString())
            __p((merged["c"]).toString())
        
__check("1\n2\n3")
}
