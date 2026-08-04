// vybe-test: kotlin/collections_sequences/test_sequence_from_list_is_lazy_until_terminal
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
            var built = 0
            val source = listOf(1, 2, 3)
            val seq = source.asSequence().onEach { built += 1 }
            __p(("before").toString())
            __p((seq.count()).toString())
            __p(("after").toString())
            __p((built).toString())
        
__check("before\n3\nafter\n3")
}
