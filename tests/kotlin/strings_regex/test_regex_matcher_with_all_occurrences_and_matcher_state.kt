// vybe-test: kotlin/strings_regex/test_regex_matcher_with_all_occurrences_and_matcher_state
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

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
            val pattern = Regex("\\b\\w+\\b")
            val matcher = pattern.toPattern().matcher("one two three")
            var matches = ""
            while (matcher.find()) {
                matches += matcher.group()
                matches += ":"
                matches += matcher.start().toString()
                matches += "-"
                matches += matcher.end().toString()
                matches += ";"
            }
            __p((matches).toString())
        
__check("one:0-3;two:4-7;three:8-13;")
}
