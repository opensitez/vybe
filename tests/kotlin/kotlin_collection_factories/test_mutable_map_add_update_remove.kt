// vybe-test: kotlin/kotlin_collection_factories/test_mutable_map_add_update_remove
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

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
            val value = mutableMapOf("x" to 1)
            value["y"] = 2
            value["x"] = 9
            value.remove("y")
            __p((value["x"]).toString())
            __p((value.containsKey("y")).toString())
            __p((value.size).toString())
        
__check("9\nfalse\n1")
}
