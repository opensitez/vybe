// vybe-test: kotlin/collection_projection_views/test_map_values_mutable_list_backed_view
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
            val map = linkedMapOf("x" to 1, "y" to 2)
            val values = map.values
            __p((values.sum()).toString())
            map["x"] = 9
            __p((values.joinToString(",")).toString())
            val copied = values.toMutableList()
            copied.remove(2)
            __p((values.size).toString())
            __p((copied.size).toString())
        
__check("3\n9,2\n2\n1")
}
