// vybe-test: kotlin/local_functions/test_local_function_with_multiple_return_paths
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

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
            fun classify(v: Int): String {
                if (v < 0) return "neg"
                if (v == 0) return "zero"
                return "pos"
            }
            __p((classify(-1)).toString())
            __p((classify(0)).toString())
            __p((classify(1)).toString())
        
__check("neg\nzero\npos")
}
