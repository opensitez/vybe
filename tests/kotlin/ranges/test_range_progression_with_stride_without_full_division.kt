// vybe-test: kotlin/ranges/test_range_progression_with_stride_without_full_division
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

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
            var values = ""
            for (value in 0..11 step 4) {
                values += value.toString()
                if (value > 8) {
                    values += "|"
                }
            }
            __p((values).toString())
            __p((10 in (0..11 step 4)).toString())
            __p((8 in (0..11 step 4)).toString())
        
__check("048|\nfalse\ntrue")
}
