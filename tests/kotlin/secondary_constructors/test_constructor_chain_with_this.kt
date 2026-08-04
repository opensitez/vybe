// vybe-test: kotlin/secondary_constructors/test_constructor_chain_with_this
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Margin {
            val top: Int
            val right: Int
            val bottom: Int
            val left: Int

            constructor(all: Int) : this(all, all, all, all) {
                __p(("all").toString())
            }

            constructor(top: Int, right: Int, bottom: Int, left: Int) {
                this.top = top
                this.right = right
                this.bottom = bottom
                this.left = left
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
            val m = Margin(7)
            __p((m.top).toString())
            __p((m.right).toString())
            __p((m.bottom).toString())
            __p((m.left).toString())
        
__check("all\n7\n7\n7\n7")
}
