// vybe-test: kotlin/operators/test_range_and_contains_for_custom_window
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Window(val low: Int, val high: Int) {
            operator fun contains(value: Int): Boolean {
                return value >= low && value <= high
            }

            operator fun rangeTo(other: Int): IntRange {
                return low..other
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
            val window = Window(1, 4)
            __p((2 in window).toString())
            __p((6 in window).toString())
            val span = window..5
            var total = 0
            for (value in span) {
                total += value
            }
            __p((total).toString())
        
__check("true\nfalse\n15")
}
