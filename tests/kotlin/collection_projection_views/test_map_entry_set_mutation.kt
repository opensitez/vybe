// vybe-test: kotlin/collection_projection_views/test_map_entry_set_mutation
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

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
            val map = linkedMapOf("a" to 1, "b" to 2)
            val entries = map.entries
            val it = entries.iterator()
            while (it.hasNext()) {
                val e = it.next()
                if (e.key == "a") {
                    it.remove()
                }
            }
            __p((map.size).toString())
            __p((map.containsKey("a")).toString())
            __p((map["b"]).toString())
        
__check("1\nfalse\n2")
}
