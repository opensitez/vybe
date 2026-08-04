// vybe-test: kotlin/kotlin_visibility_keywords/test_private_and_public_property_access
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_keywords.rs

open class Base {
            private val secret = "secret"
            public val shown = "shown"
            protected open val inherited = "inherited"
        }

        class Child : Base() {
            override val inherited: String = "childInherited"
            fun exposeInherited(): String {
                return inherited
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
            val b = Child()
            __p((b.shown).toString())
            __p((b.exposeInherited()).toString())
        
__check("shown\nchildInherited")
}
