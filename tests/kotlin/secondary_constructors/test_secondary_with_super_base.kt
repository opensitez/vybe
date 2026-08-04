// vybe-test: kotlin/secondary_constructors/test_secondary_with_super_base
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

open class P(val id: Int)

        class C : P {
            val tag: Int

            constructor() : super(1) {
                this.tag = 2
            }

            constructor(multiplier: Int) : super(multiplier) {
                this.tag = multiplier * 2
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
            __p((C().id).toString())
            __p((C(3).id).toString())
            __p((C(3).tag).toString())
        
__check("1\n3\n6")
}
