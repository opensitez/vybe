// vybe-test: kotlin/collections_maps_ops/test_map_entries_iteration_order_in_linked_map
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
            val map = linkedMapOf("first" to 1, "second" to 2, "third" to 3)
            var keys = ""
            var sum = 0
            for (entry in map.entries) {
                keys += entry.key
                sum += entry.value
            }
            __p((keys).toString())
            __p((sum).toString())
        
__check("firstsecondthird\n6")
}
