// vybe-test: kotlin/strings_regex/test_regex_find_all_uses_capturing_groups
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
            val pattern = Regex("(\\w+)(\\d)")
            val matches = pattern.findAll("a1 b22 c3")
            var output = ""
            for (item in matches) {
                output += item.groupValues[1]
                output += "-"
                output += item.groupValues[2]
                output += ";"
            }
            __p((output).toString())
        
__check("a-1;b2-2;c-3;")
}
