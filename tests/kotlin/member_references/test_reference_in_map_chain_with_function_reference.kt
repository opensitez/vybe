// vybe-test: kotlin/member_references/test_reference_in_map_chain_with_function_reference
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Item(val value: Int)

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
            val items = listOf(Item(3), Item(7), Item(9))
            val refs = items.map(Item::value).map { it * 2 }
            __p((refs.joinToString("|")).toString())
        
__check("6|14|18")
}
