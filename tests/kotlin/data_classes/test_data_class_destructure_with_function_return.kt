// vybe-test: kotlin/data_classes/test_data_class_destructure_with_function_return
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Record(val code: Int, val weight: Int)

        fun split(): Record {
            return Record(4, 5)
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
            val (code, weight) = split()
            __p((code).toString())
            __p((weight).toString())
            __p((code + weight).toString())
        
__check("4\n5\n9")
}
