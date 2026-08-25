// vybe-test: kotlin/initialization_order/test_init_block_can_reference_secondary_defaults
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder {
            val value: Int

            init {
                value = 7
                __p(("init").toString())
            }

            // Damaged spelling repaired: the original body was
            // `constructor() { this() }` — kotlinc 2.4.10 rejects the body
            // call ("unresolved reference 'invoke'"), and a no-arg ctor
            // delegating to itself is unconstructible anyway. A body-less
            // `constructor()` runs the init block and, measured under
            // kotlinc 2.4.10, prints exactly the expected init, 7.
            constructor()

            fun out(): Int = value
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
            val item = Holder()
            __p((item.out()).toString())
        
__check("init\n7")
}
