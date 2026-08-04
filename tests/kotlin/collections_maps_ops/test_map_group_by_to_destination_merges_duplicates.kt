// vybe-test: kotlin/collections_maps_ops/test_map_group_by_to_destination_merges_duplicates
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

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
            val words = listOf("a", "bb", "cc", "ddd", "eee")
            val groups = mutableMapOf<Int, MutableList<String>>()
            words.groupByTo(groups, { it.length })
            __p((groups[1]?.size).toString())
            __p((groups[2]?.size).toString())
            __p((groups[3]?.size).toString())
        
__check("1\n2\n2")
}
