// vybe-test: kotlin/inheritance_dispatch/test_virtual_property_dispatch_in_inheritance_chain
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open val value: Int = 1
            open fun total(): Int = value + 1
        }

        class Child : Base() {
            override val value: Int = 3
            override fun total(): Int = value + 2
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
            val item: Base = Child()
            __p((item.value).toString())
            __p((item.total()).toString())
        
__check("3\n5")
}
