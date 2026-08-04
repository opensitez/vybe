// vybe-test: kotlin/enums/test_enum_array_iteration_sum
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Digit { D0, D1, D2, D3 }
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

fun main() { var n = 0
for (d in arrayOf(Digit.D0, Digit.D1, Digit.D2, Digit.D3)) { n += d }
__p((n).toString()) 
__check("6")
}
