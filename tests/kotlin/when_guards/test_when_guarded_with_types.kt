// vybe-test: kotlin/when_guards/test_when_guarded_with_types
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun label(v: Any): String = when {
            v is String && v.isEmpty() -> "empty"
            v is String -> "str"
            v is Int && v > 10 -> "big-int"
            v is Int -> "int"
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
            __p((label("")).toString())
            __p((label("x")).toString())
            __p((label(11)).toString())
            __p((label(5)).toString())
        
__check("empty\nstr\nbig-int\nint")
}
