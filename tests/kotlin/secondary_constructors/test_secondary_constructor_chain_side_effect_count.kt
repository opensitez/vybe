// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_chain_side_effect_count
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class SequenceTracker {
            val value: Int

            constructor() {
                __p(("base").toString())
                this.value = 0
            }

            constructor(start: Int) : this() {
                __p(("fromStart").toString())
                this.value = start
            }

            constructor(start: Int, step: Int) : this(start) {
                __p(("fromStep").toString())
                this.value = start + step
            }
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
            __p((SequenceTracker().value).toString())
            __p((SequenceTracker(3).value).toString())
            __p((SequenceTracker(3, 4).value).toString())
        
__check("base\n0\nfromStart\n3\nfromStep\n7")
}
