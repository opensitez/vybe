// vybe-test: kotlin/scoping_functions/test_apply_with_nested_with_reads_its_properties
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Item {
            val id = 1
            val prefix = "item"
        }

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
            val text = Item().apply {
                id.toString()
            }.let { it.id }
            __p((text).toString())
            val withText = with(Item()) { "$prefix-$id" }
            __p((withText).toString())
        
__check("1\nitem-1")
}
