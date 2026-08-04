// vybe-test: kotlin/kotlin_lazy_delegates/test_lazy_after_mutation_from_other_property
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

class Holder {
            var seed = 1
            val value by lazy { seed * 10 }
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
            val h = Holder()
            h.seed = 3
            __p((h.value).toString())
            h.seed = 9
            __p((h.value).toString())
        
__check("30\n30")
}
