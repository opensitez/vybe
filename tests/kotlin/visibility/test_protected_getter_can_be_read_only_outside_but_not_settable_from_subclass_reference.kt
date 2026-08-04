// vybe-test: kotlin/visibility/test_protected_getter_can_be_read_only_outside_but_not_settable_from_subclass_reference
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

open class Base {
            protected var value: Int = 1
        }

        class Child : Base() {
            var view: Int
                get() = value
                set(next) { value = next }
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
            val child = Child()
            child.view = 10
            __p((child.view).toString())
        
__check("10")
}
