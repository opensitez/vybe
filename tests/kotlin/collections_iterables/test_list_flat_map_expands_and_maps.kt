// vybe-test: kotlin/collections_iterables/test_list_flat_map_expands_and_maps
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

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
            val groups = listOf(
                listOf(1, 2),
                listOf(3, 4)
            )
            val expanded = groups.flatMap { it }
            __p((expanded.joinToString(",")).toString())
            val mapped = groups.flatMap { inner -> inner.map { it * 10 } }
            __p((mapped.joinToString(",")).toString())
        
__check("1,2,3,4\n10,20,30,40")
}
