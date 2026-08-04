// vybe-test: kotlin/operators/test_operator_with_generic_type
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Box<T>(private val value: T) {
            operator fun plus(other: Box<T>): String {
                return this.value.toString() + other.value.toString()
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
            __p((Box("a") + Box("b")).toString())
            __p((Box(1) + Box(2)).toString())
        
__check("ab\n12")
}
