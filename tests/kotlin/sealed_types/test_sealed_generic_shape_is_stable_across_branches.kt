// vybe-test: kotlin/sealed_types/test_sealed_generic_shape_is_stable_across_branches
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Value {
            class Text(val value: String) : Value()
            class Count(val value: Int) : Value()
        }

        fun normalize(value: Value): String {
            return when (value) {
                is Value.Text -> value.value
                is Value.Count -> value.value.toString()
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
            val list: List<Value> = listOf(Value.Text("x"), Value.Count(3))
            __p((normalize(list[0])).toString())
            __p((normalize(list[1])).toString())
        
__check("x\n3")
}
