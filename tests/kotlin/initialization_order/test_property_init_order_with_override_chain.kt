// vybe-test: kotlin/initialization_order/test_property_init_order_with_override_chain
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class Base {
            open val base = 1
            init {
                __p((base).toString())
            }
        }

        class Child : Base() {
            override val base = 4
            init {
                __p((base).toString())
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
            Child()

// Damaged expectation repaired: the original wanted "4\n4", but Base's init
// block reads the OVERRIDDEN getter while Child's backing field still holds
// the Int default — measured under kotlinc 2.4.10 this program prints 0
// then 4 (the classic leaking-`this` initialization-order gotcha).
__check("0\n4")
}
