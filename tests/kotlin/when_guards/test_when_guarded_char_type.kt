// vybe-test: kotlin/when_guards/test_when_guarded_char_type
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun toCategory(c: Char): String = when {
            c == 'x' || c == 'y' -> "xy"
            c in 'a'..'f' -> "alpha"
            c.isDigit() -> "digit"
            else -> "other"
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
            __p((toCategory('x')).toString())
            __p((toCategory('b')).toString())
            __p((toCategory('7')).toString())
        
__check("xy\nalpha\ndigit")
}
