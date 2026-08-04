// vybe-test: kotlin/strings/test_substring_before_last_and_after_last_boundaries
// origin: languages/kotlin/tests/kotlin/test_strings.rs

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
            val value = "a/b/c"
            __p((value.substringBeforeLast("/")).toString())
            __p((value.substringAfterLast("/")).toString())
            __p(("nodelim".substringAfterLast("/", "none")).toString())
            __p(("x/y/".substringAfterLast("/")).toString())
        
__check("a/b\nc\nnone\n")
}
