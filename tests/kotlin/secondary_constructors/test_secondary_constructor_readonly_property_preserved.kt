// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_readonly_property_preserved
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Metric {
            val total: Int
            var tag: String

            constructor(base: Int) {
                this.total = base
                this.tag = "base"
            }

            constructor(base: Int, tag: String) : this(base) {
                this.tag = tag
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
            val one = Metric(3)
            val two = Metric(5, "custom")
            __p((one.total).toString())
            __p((one.tag).toString())
            __p((two.total).toString())
            __p((two.tag).toString())
        
__check("3\nbase\n5\ncustom")
}
