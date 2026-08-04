// vybe-test: kotlin/kotlin_initializer_blocks/test_init_block_with_arguments
// origin: languages/kotlin/tests/kotlin/test_kotlin_initializer_blocks.rs

class Calc(val base: Int) {
            val offset = base + 1

            init {
                __p(((base * offset).toString()).toString())
            }

            init {
                __p(((offset - base).toString()).toString())
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
            Calc(3)
        
__check("12\n1")
}
