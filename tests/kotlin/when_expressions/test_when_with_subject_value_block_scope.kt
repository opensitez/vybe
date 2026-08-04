// vybe-test: kotlin/when_expressions/test_when_with_subject_value_block_scope
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun render(level: Int): String {
            return when (level) {
                in 0..9 -> {
                    val label = "low"
                    label + ":" + level
                }
                in 10..19 -> {
                    val offset = level - 10
                    "mid:" + offset
                }
                else -> {
                    val doubled = level * 2
                    "high:" + doubled
                }
            }
        }

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
            __p((render(4)).toString())
            __p((render(13)).toString())
            __p((render(30)).toString())
        
__check("low:4\nmid:3\nhigh:60")
}
