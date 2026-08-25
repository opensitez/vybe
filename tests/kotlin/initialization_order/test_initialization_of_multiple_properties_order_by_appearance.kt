// vybe-test: kotlin/initialization_order/test_initialization_of_multiple_properties_order_by_appearance
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

// Damaged spelling repaired: the original declared `val b = a + c` BEFORE
// `c`, and kotlinc 2.4.10 rejects the forward reference — "variable 'c' must
// be initialized". Declaring `c` before `b` keeps the test's point (properties
// initialize in order of appearance) and, measured under kotlinc 2.4.10,
// prints exactly the expected 1, 4, 3.
class Holder {
            val a = 1
            val c = 3
            val b = a + c

            init {
                __p((a).toString())
                __p((b).toString())
                __p((c).toString())
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
            Holder()
        
__check("1\n4\n3")
}
