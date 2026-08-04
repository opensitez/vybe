// vybe-test: kotlin/inheritance_dispatch/test_dispatch_in_higher_order_function_argument
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Labelled {
            open fun text(): String = "base"
        }

        class Dynamic : Labelled() {
            override fun text(): String = "dynamic"
        }

        fun mapLabel(value: Labelled, render: (Labelled) -> String): String {
            return render(value)
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
            val item: Labelled = Dynamic()
            __p((mapLabel(item) { it.text() }).toString())
            __p((mapLabel(item) { target -> "[" + target.text() + "]" }).toString())
        
__check("dynamic\n[dynamic]")
}
