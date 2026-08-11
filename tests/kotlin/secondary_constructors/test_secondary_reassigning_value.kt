// vybe-test: kotlin/secondary_constructors/test_secondary_reassigning_value
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Counter {
            var value: Int

            constructor() {
                this.value = 1
            }

            constructor(value: Int, double: Boolean) : this() {
                if (double) {
                    this.value = value * 2
                } else {
                    this.value = value
                }
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
            __p((Counter(3, true).value).toString())
            __p((Counter(3, false).value).toString())
        
__check("6\n3")
}
