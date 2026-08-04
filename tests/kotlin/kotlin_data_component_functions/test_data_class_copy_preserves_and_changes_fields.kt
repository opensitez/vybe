// vybe-test: kotlin/kotlin_data_component_functions/test_data_class_copy_preserves_and_changes_fields
// origin: languages/kotlin/tests/kotlin/test_kotlin_data_component_functions.rs

data class Point(val x: Int, val y: Int)

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
            val point = Point(2, 3)
            val shifted = point.copy(y = 9)
            __p((point).toString())
            __p((shifted).toString())
        
__check("Point(x=2, y=3)\nPoint(x=2, y=9)")
}
