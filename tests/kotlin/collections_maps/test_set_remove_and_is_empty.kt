// vybe-test: kotlin/collections_maps/test_set_remove_and_is_empty
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
            val ids = mutableSetOf(1, 2, 3)
            __p((ids.remove(2)).toString())
            __p((ids.remove(4)).toString())
            __p((ids.isNotEmpty()).toString())
            ids.remove(1)
            ids.remove(3)
            __p((ids.isEmpty()).toString())
            __p((ids.size).toString())
        
__check("true\nfalse\ntrue\ntrue\n0")
}
