// vybe-test: kotlin/collections_sequences/test_sequence_map_indexed_keeps_index_contract
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
            val labeled = listOf("x", "y", "z").asSequence()
                .mapIndexed { index, value -> "$index:$value" }
                .toList()
            __p((labeled.joinToString(",")).toString())
        
__check("0:x,1:y,2:z")
}
