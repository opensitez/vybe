// vybe-test: kotlin/kotlin_progressions/test_progression_take_drop
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

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
            val r = (1..10)
            val first = r.take(4)
            val remain = r.drop(4)
            __p((first.joinToString(",")).toString())
            __p((remain.take(3).joinToString(",")).toString())
        
__check("[1, 2, 3, 4]\n[5, 6, 7]")
}
