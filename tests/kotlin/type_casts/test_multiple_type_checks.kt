// vybe-test: kotlin/type_casts/test_multiple_type_checks
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun check(value: Any) {
            if (value is String) {
                __p(("string").toString())
            } else if (value !is Int) {
                __p(("not int").toString())
            } else {
                __p(("int").toString())
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
            check("x")
            check(3)
            check(true)
        
__check("string\nint\nnot int")
}
