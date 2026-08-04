// vybe-test: kotlin/generics/test_generic_function_returning_array_and_size
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> toArray(left: T, right: T): Array<T> {
            return arrayOf(left, right)
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
            val numbers = toArray(2, 3)
            val words = toArray("a", "b")
            __p((numbers.size).toString())
            __p((words.size).toString())
            __p((numbers[1] + words[1]).toString())
        
__check("2\n2\n3b")
}
