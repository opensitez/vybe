// vybe-test: kotlin/secondary_constructors/test_constructor_three_levels
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Scale {
            val unit: Int

            constructor() {
                this.unit = 1
            }

            constructor(value: Int) : this() {
                __p(("scaled").toString())
            }

            constructor(value: Int, factor: Int) : this(value) {
                __p((value * factor).toString())
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
            Scale()
            Scale(4)
            Scale(5, 2)
        
__check("scaled\nscaled\n10")
}
