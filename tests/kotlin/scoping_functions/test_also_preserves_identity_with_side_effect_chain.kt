// vybe-test: kotlin/scoping_functions/test_also_preserves_identity_with_side_effect_chain
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

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
            val first = mutableListOf(1)
            val second = first
                .also { it.add(2) }
                .also { it.add(3) }
            __p((first === second).toString())
            __p((second.joinToString("|")).toString())
        
__check("true\n1|2|3")
}
