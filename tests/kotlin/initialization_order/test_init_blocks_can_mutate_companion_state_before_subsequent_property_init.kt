// vybe-test: kotlin/initialization_order/test_init_blocks_can_mutate_companion_state_before_subsequent_property_init
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

var globalValue = 1

        class Holder {
            val first = globalValue

            init {
                globalValue = 10
            }

            val second = globalValue * 2

            init {
                __p((first).toString())
                __p((second).toString())
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
            Holder()
            __p((globalValue).toString())
        
__check("1\n20\n10\n20\n10")
}
