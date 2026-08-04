// vybe-test: kotlin/classes/test_class_with_constructor_sharing
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class PairNode {
            val left: Int
            val right: Int

            constructor(left: Int, right: Int) {
                this.left = left
                this.right = right
            }

            constructor(value: Int) : this(value, value) {
                __p(("copy").toString())
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
            val p1 = PairNode(4)
            val p2 = PairNode(1, 3)
            __p((p1.left).toString())
            __p((p1.right).toString())
            __p((p2.left).toString())
            __p((p2.right).toString())
        
__check("copy\n4\n4\n1\n3")
}
