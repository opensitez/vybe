// vybe-test: kotlin/kotlin_regex_advanced/test_regex_match_entire_line_with_options_combo
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

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
            val pattern = Regex("^\n*OK\\?$")
            val withComments = Regex("^\n*OK\\?$")
            __p((pattern.matches("OK?")).toString())
            __p((withComments.matches("\n\nOK?")).toString())
            val withOption = Regex("^\n*OK\\?$", RegexOption.MULTILINE)
            __p((withOption.matches("line1\nOK?")).toString())
        
__check("true\nfalse\nfalse")
}
