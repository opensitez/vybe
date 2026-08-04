// vybe-test: kotlin/kotlin_pairs_triples/test_triple_destructuring_in_function_return
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun make(): Triple<String, Int, Boolean> {
            return Triple("x", 4, true)
        }

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
            val (k, n, b) = make()
            __p((k).toString())
            __p((n).toString())
            __p((b).toString())
        
__check("x\n4\ntrue")
}
