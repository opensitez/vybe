// vybe-test: kotlin/when_expressions/test_when_subject_capture_with_is_and_casts
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

interface Shape { fun kind(): String }
        class Dot : Shape {
            override fun kind(): String = "dot"
        }
        class Box(val size: Int) : Shape {
            override fun kind(): String = "box:" + size
        }

        fun describe(shape: Shape): String {
            return when (shape) {
                is Dot -> "dot"
                is Box -> "box=" + shape.size
                else -> "unknown"
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
            __p((describe(Dot())).toString())
            __p((describe(Box(7))).toString())
        
__check("dot\nbox=7")
}
