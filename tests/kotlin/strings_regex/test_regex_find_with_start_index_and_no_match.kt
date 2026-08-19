// vybe-test: kotlin/strings_regex/test_regex_find_with_start_index_and_no_match
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
            val pattern = Regex("\\d+")
            __p((pattern.find("abc123", 1)?.value ?: "none").toString())
            __p((pattern.find("abc123", 4)?.value ?: "none").toString())
            __p((pattern.find("abc", 3) == null).toString())
            var beyond = "no throw"
            try {
                pattern.find("abc", 5)
            } catch (e: IndexOutOfBoundsException) {
                beyond = "threw"
            }
            __p((beyond).toString())
        
__check("123\n23\ntrue\nthrew")
}
