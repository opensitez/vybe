// vybe-test: kotlin/collections_sequences/test_sequence_flat_map_nested_lists
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

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
            val seq = listOf(listOf(1, 2), listOf(3, 4)).asSequence().flatMap { it.asSequence() }
            __p((seq.sum()).toString())
            __p((seq.toList().joinToString(",")).toString())
        
__check("10\n1,2,3,4")
}
