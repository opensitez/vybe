// vybe-test: kotlin/initialization_order/test_init_uses_global_counter_and_reuses_updated_state
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

var stamp = 0

        fun next_stamp(): Int {
            stamp += 1
            return stamp
        }

        class Holder {
            val first = next_stamp()
            val second = first + 10
            init {
                __p((first).toString())
                __p((second).toString())
            }
            val third = next_stamp()
            init {
                __p((third).toString())
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
            __p((stamp).toString())
        
__check("1\n11\n2\n3\n13\n4\n4")
}
