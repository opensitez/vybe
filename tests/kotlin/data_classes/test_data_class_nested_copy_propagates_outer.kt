// vybe-test: kotlin/data_classes/test_data_class_nested_copy_propagates_outer
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Child(val value: Int)
        data class Parent(val child: Child, val tag: String)

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
            val p1 = Parent(Child(1), "x")
            val p2 = p1.copy(child = Child(9))
            __p((p1.child.value).toString())
            __p((p2.child.value).toString())
            __p((p2.tag).toString())
        
__check("1\n9\nx")
}
