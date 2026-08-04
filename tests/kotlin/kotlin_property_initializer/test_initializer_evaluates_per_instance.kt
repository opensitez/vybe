// vybe-test: kotlin/kotlin_property_initializer/test_initializer_evaluates_per_instance
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_initializer.rs

var marker = 0

        fun step(): Int {
            marker = marker + 10
            return marker
        }

        class Token {
            val value = step()
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
            val first = Token()
            val second = Token()
            __p((first.value).toString())
            __p((second.value).toString())
        
__check("10\n20")
}
