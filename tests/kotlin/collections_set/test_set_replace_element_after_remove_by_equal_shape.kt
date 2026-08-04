// vybe-test: kotlin/collections_set/test_set_replace_element_after_remove_by_equal_shape
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

data class Box(val id: Int)

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
            val values = mutableSetOf(Box(1), Box(2))
            values.remove(Box(1))
            values.add(Box(3))
            __p((values.size).toString())
            __p((values.contains(Box(2))).toString())
            __p((values.contains(Box(1))).toString())
            __p((values.contains(Box(3))).toString())
        
__check("2\ntrue\nfalse\ntrue")
}
