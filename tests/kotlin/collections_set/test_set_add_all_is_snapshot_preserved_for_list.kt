// vybe-test: kotlin/collections_set/test_set_add_all_is_snapshot_preserved_for_list
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

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
            val source = setOf(1, 2, 3)
            val copied = source.toMutableSet()
            copied.addAll(listOf(3, 4, 5))
            __p((source.toString()).toString())
            __p((copied.toString()).toString())
            __p((source.contains(5)).toString())
            __p((copied.contains(5)).toString())
        
__check("[1, 2, 3]\n[1, 2, 3, 4, 5]\nfalse\ntrue")
}
