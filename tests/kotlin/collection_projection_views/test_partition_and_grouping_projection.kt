// vybe-test: kotlin/collection_projection_views/test_partition_and_grouping_projection
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
            val values = listOf(1, 2, 3, 4, 5)
            val (even, odd) = values.partition { it % 2 == 0 }
            __p((even.joinToString(",")).toString())
            __p((odd.joinToString(",")).toString())
            val byMod = values.groupBy { it % 2 }
            __p((byMod[0]!!.joinToString(",")).toString())
            __p((byMod[1]!!.joinToString(",")).toString())
        
__check("2,4\n1,3,5\n2,4\n1,3,5")
}
