// vybe-test: kotlin/kotlin_accessor_customization/test_custom_getter
// origin: languages/kotlin/tests/kotlin/test_kotlin_accessor_customization.rs

class Score {
            private var raw = 0

            var value: Int
                get() = raw * 2
                set(v) { raw = if (v < 0) 0 else v }
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
            val s = Score()
            s.value = 3
            __p((s.value).toString())
            s.value = -4
            __p((s.value).toString())
        
__check("6\n0")
}
