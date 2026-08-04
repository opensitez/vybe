// vybe-test: kotlin/strings_regex/test_regex_replace_with_counted_callback
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
            var index = 0
            val pattern = Regex("\\d")
            val output = pattern.replace("a1b2c3") { match ->
                val value = "${index}:${match.value}"
                index += 1
                value
            }
            __p((output).toString())
            __p((index).toString())
        
__check("a0:1b1:2c2:3\n3")
}
