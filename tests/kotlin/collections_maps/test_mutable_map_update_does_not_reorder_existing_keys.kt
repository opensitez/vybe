// vybe-test: kotlin/collections_maps/test_mutable_map_update_does_not_reorder_existing_keys
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
            val state = linkedMapOf("first" to 1, "second" to 2)
            state["first"] = 9
            state["first"] = 11
            var keys = ""
            for ((key, _) in state) {
                keys += key
            }
            __p((keys).toString())
            __p((state["first"]).toString())
            __p((state.size).toString())
        
__check("firstsecond\n11\n2")
}
