// vybe-test: kotlin/interfaces/test_interface_array_dispatch_across_types
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Token {
            fun kind(): String
        }

        class Alpha : Token {
            override fun kind(): String = "alpha"
        }

        class Beta : Token {
            override fun kind(): String = "beta"
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
            val tokens: Array<Token> = arrayOf(Alpha(), Beta(), Alpha())
            var alpha = 0
            var beta = 0
            for (token in tokens) {
                when (token.kind()) {
                    "alpha" -> alpha += 1
                    "beta" -> beta += 1
                }
            }
            __p((alpha).toString())
            __p((beta).toString())
        
__check("2\n1")
}
