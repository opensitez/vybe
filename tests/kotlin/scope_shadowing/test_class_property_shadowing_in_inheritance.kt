// vybe-test: kotlin/scope_shadowing/test_class_property_shadowing_in_inheritance
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

open class Parent {
            val value = "parent"
        }
        class Child : Parent() {
            val value = "child"
            fun reveal() = super.value + ":" + value
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
            val c = Child()
            __p((c.value).toString())
            __p((c.reveal()).toString())
        
__check("child\nparent:child")
}
