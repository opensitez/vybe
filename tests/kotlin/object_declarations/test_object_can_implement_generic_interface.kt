// vybe-test: kotlin/object_declarations/test_object_can_implement_generic_interface
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

interface Transformer<T, U> {
            fun map(value: T): U
        }

        object IntToText : Transformer<Int, String> {
            override fun map(value: Int): String = "v" + value
        }

        fun emit(transformer: Transformer<Int, String>, value: Int): String {
            return transformer.map(value)
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
            __p((emit(IntToText, 3)).toString())
        
__check("v3")
}
