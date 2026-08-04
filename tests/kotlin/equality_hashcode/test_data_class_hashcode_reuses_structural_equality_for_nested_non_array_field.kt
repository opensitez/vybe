// vybe-test: kotlin/equality_hashcode/test_data_class_hashcode_reuses_structural_equality_for_nested_non_array_field
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Child(val token: String)
        data class Parent(val child: Child)

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
            val left = Parent(Child("x"))
            val right = Parent(Child("x"))
            __p((left == right).toString())
            __p((left.hashCode() == right.hashCode()).toString())
        
__check("true\ntrue")
}
