// vybe-test: kotlin/secondary_constructors/test_constructor_property_assignment
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Rectangle {
            val width: Int
            val height: Int

            constructor(width: Int, height: Int) {
                this.width = width
                this.height = height
            }

            constructor(size: Int) : this(size, size) {
                __p(("square").toString())
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
            val r1 = Rectangle(2, 3)
            val r2 = Rectangle(4)
            __p((r1.width).toString())
            __p((r1.height).toString())
            __p((r2.width).toString())
            __p((r2.height).toString())
        
__check("square\n2\n3\n4\n4")
}
