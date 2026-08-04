// vybe-test: kotlin/equality_hashcode/test_equality_is_reflexive_and_symmetric_for_data_classes
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Point(val x: Int, val y: Int)

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
            val first = Point(1, 2)
            val second = Point(1, 2)
            __p((first == first).toString())
            __p((first == second).toString())
            __p((second == first).toString())
        
__check("true\ntrue\ntrue")
}
