// vybe-test: kotlin/collections_maps_ops/test_map_get_or_put_without_recomputing_for_present_key
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
            var computed = 0
            val map = mutableMapOf("present" to 1)
            val value1 = map.getOrPut("present") {
                computed += 1
                99
            }
            val value2 = map.getOrPut("missing") {
                computed += 1
                77
            }
            __p((value1).toString())
            __p((value2).toString())
            __p((computed).toString())
        
__check("1\n77\n1")
}
