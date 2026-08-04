// vybe-test: kotlin/initialization_order/test_secondary_constructor_delegation_preserves_base_initialization
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class Base(prefix: Int) {
            val value = prefix
            init {
                __p((value).toString())
            }
        }

        class Leaf : Base {
            val label: Int
            constructor() : this(2)
            constructor(seed: Int) : super(seed) {
                label = seed * 10
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
            __p((Leaf().label).toString())
        
__check("2\n20")
}
