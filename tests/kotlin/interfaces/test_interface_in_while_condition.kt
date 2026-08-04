// vybe-test: kotlin/interfaces/test_interface_in_while_condition
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Marker { fun hit(): Boolean }
class Yes: Marker { override fun hit(): Boolean = true }
class No: Marker { override fun hit(): Boolean = false }
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

fun main() { var score = 0
for (m in arrayOf(Yes(), No())) { if (m.hit()) score += 1 }
__p((score).toString()) 
__check("1")
}
