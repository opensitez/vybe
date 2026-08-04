// vybe-test: kotlin/function_overloads/test_overload_in_generic_class_method
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

class Box {
            fun size(value: Int): String = "int"
            fun size(value: String): String = "str"
            fun size(value: List<Int>): String = "list"
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
            val b = Box()
            __p((b.size(4)).toString())
            __p((b.size("a")).toString())
            __p((b.size(listOf(1))).toString())
        
__check("int\nstr\nlist")
}
