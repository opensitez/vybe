// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_with_custom_tie_breaker
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

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
            data class Item(val first: Int, val second: String)
            val values = listOf(Item(2, "b"), Item(1, "c"), Item(2, "a"))
            val out = values.sortedWith(compareBy<Item> { it.first }.thenBy { it.second })
            __p((out.joinToString(",") { "${it.first}${it.second}" }).toString())
        
__check("1c,2a,2b")
}
