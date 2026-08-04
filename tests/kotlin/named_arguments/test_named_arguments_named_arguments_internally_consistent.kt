// vybe-test: kotlin/named_arguments/test_named_arguments_named_arguments_internally_consistent
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun total(one: Int = 1, two: Int = 2, three: Int = 3): Int {
            return one + two + three
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
            __p((total()).toString())
            __p((total(two = 10)).toString())
            __p((total(three = 7, one = 1, two = 2)).toString())
        
__check("6\n12\n10")
}
