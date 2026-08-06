// vybe-test: kotlin/strings/test_lines_and_trim_with_blank_lines
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
            val value = "a\n\nb\n"
            val raw = value.lines()
            __p((raw.size).toString())
            __p((raw[1]).toString())
            __p((raw[2].isEmpty()).toString())
            __p((value.lines().filter { it.isNotEmpty() }.joinToString("|")).toString())
        
__check("4\n\nfalse\na|b")
}
