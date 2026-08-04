// vybe-test: kotlin/data_class_copying/test_data_class_copy_multiple_instances
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Row(val id: Int)
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
            val first = Row(1)
            val second = first.copy(2)
            val third = second.copy(3)
            __p((first.id).toString())
            __p((second.id).toString())
            __p((third.id).toString())
        
__check("1\n2\n3")
}
