// vybe-test: kotlin/kotlin_nested_labels/test_nested_labeled_while
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_labels.rs

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
            var i = 0
            outer@ while (i < 3) {
                var j = 0
                while (j < 2) {
                    if (i == 1 && j == 0) {
                        j = j + 1
                        i = i + 1
                        continue@outer
                    }
                    __p((i.toString() + ":" + j.toString()).toString())
                    j = j + 1
                }
                i = i + 1
            }
        
__check("0:0\n0:1\n2:0\n2:1")
}
