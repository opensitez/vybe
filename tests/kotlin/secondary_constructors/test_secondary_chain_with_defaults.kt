// vybe-test: kotlin/secondary_constructors/test_secondary_chain_with_defaults
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Packet {
            val x: Int

            constructor() {
                this.x = 0
            }

            constructor(v: Int) : this() {
                this.x = v
            }

            constructor(v: Int, d: Int, e: Int) : this(v) {
                this.x = v + d + e
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
            val p = Packet(2, 3, 4)
            __p((p.x).toString())
        
__check("9")
}
