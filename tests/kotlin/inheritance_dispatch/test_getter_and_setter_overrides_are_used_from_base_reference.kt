// vybe-test: kotlin/inheritance_dispatch/test_getter_and_setter_overrides_are_used_from_base_reference
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open var value: Int = 0
        }

        class Child : Base() {
            private var storage = 10

            override var value: Int
                get() = storage
                set(new_value) {
                    storage = new_value + 1
                }
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
            val base: Base = Child()
            base.value = 7
            __p((base.value).toString())
            __p(((base as Child).value).toString())
        
__check("8\n8")
}
