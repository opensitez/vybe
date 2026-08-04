// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_index_even_odd_count
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

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
            val grouped = (1..7).withIndex().groupBy { it.index % 2 }
            __p((grouped[0]!!.size).toString())
            __p((grouped[1]!!.size).toString())
        
__check("4\n3")
}
