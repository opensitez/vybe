// vybe-test: kotlin/class_delegation/test_simple_class_delegation_forwards_call
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Greeter { fun hello(): String }

        class Base(private val name: String) : Greeter {
            override fun hello() = "hello:$name"
        }

        class Wrapper(delegate: Greeter) : Greeter by delegate

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
            val w = Wrapper(Base("kotlin"))
            __p((w.hello()).toString())
        
__check("hello:kotlin")
}
