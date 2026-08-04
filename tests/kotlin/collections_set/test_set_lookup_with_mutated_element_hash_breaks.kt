// vybe-test: kotlin/collections_set/test_set_lookup_with_mutated_element_hash_breaks
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
            data class Item(var id: Int)
            val item = Item(1)
            val values = hashSetOf(item)
            __p((values.contains(item)).toString())
            item.id = 99
            __p((values.contains(item)).toString())
        
__check("true\nfalse")
}
