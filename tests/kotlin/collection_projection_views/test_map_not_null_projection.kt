// vybe-test: kotlin/collection_projection_views/test_map_not_null_projection
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
            val source: List<Int?> = listOf(1, null, 3, null, 5)
            __p((source.filterNotNull().joinToString(",")).toString())
            __p((source.mapNotNull { it?.plus(10) }.joinToString(",")).toString())
        
__check("1,3,5\n11,13,15")
}
