// vybe-test: kotlin/set_ordered_behavior/test_linked_hash_set_to_typed_array
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

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
            val values = linkedSetOf(7, 1, 9)
            val arr = values.toTypedArray()
            __p((arr.joinToString(",")).toString())
            val round = arr.toList().toMutableSet()
            round.add(4)
            __p((round.joinToString(",")).toString())
        
__check("7,1,9\n7,1,9,4")
}
