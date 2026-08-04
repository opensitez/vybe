// vybe-test: kotlin/list_index_api/test_list_get_or_else_and_or_null
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

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
            val values = listOf(4, 5, 6)
            __p((values.getOrElse(1) { 0 }).toString())
            __p((values.getOrElse(4) { 99 }).toString())
            __p((values.getOrNull(0) ?: -1).toString())
            __p((values.getOrNull(4) ?: -1).toString())
        
__check("5\n99\n4\n-1")
}
