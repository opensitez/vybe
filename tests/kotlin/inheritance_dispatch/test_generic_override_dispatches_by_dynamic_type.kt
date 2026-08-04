// vybe-test: kotlin/inheritance_dispatch/test_generic_override_dispatches_by_dynamic_type
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base<T> {
            open fun describe(value: T): String = "base:" + value.toString()
        }

        class Text : Base<String>() {
            override fun describe(value: String): String = "text:" + value
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
            val item: Base<String> = Text()
            __p((item.describe("ok")).toString())
            __p(((item as Text).describe("x")).toString())
        
__check("text:ok\ntext:x")
}
