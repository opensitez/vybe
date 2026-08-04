// vybe-test: kotlin/inheritance_dispatch/test_multiple_class_levels_share_virtual_method
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open fun route(): String = "base"
        }

        open class Mid : Base() {
            override fun route(): String = "mid"
        }

        class Leaf : Mid() {
            override fun route(): String = "leaf"
        }

        fun emit(route: Base): String = route.route()

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
            __p((emit(Base())).toString())
            __p((emit(Mid())).toString())
            __p((emit(Leaf())).toString())
        
__check("base\nmid\nleaf")
}
