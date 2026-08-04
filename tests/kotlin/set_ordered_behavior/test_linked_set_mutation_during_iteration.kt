// vybe-test: kotlin/set_ordered_behavior/test_linked_set_mutation_during_iteration
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
            val values = linkedSetOf(1, 2, 3)
            val outValues = StringBuilder()
            val it = values.iterator()
            while (it.hasNext()) {
                val n = it.next()
                if (n == 2) {
                    it.remove()
                }
                outValues.append(n)
            }
            __p((outValues.toString()).toString())
            __p((values.joinToString(",")).toString())
        
__check("123\n1,3")
}
