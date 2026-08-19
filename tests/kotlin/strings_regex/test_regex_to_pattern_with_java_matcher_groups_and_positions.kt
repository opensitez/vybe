// vybe-test: kotlin/strings_regex/test_regex_to_pattern_with_java_matcher_groups_and_positions
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
            val matcher = Regex("(\\w)(\\d)").toPattern().matcher("a1 b2 c3")
            var trace = ""
            while (matcher.find()) {
                trace += matcher.group(1)
                trace += matcher.group(2)
                trace += matcher.start().toString()
                trace += matcher.end().toString()
                trace += "|"
            }
            __p((trace).toString())
        
__check("a102|b235|c368|")
}
