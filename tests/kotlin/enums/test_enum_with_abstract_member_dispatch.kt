// vybe-test: kotlin/enums/test_enum_with_abstract_member_dispatch
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Operation {
            ADD {
                override fun apply(a: Int, b: Int): Int = a + b
            },
            SUBTRACT {
                override fun apply(a: Int, b: Int): Int = a - b
            },
            MULTIPLY {
                override fun apply(a: Int, b: Int): Int = a * b
            };
            abstract fun apply(a: Int, b: Int): Int
        }
// NOTE: extraction had dropped the `;` the .rs original carries after the last
// entry — Kotlin REQUIRES it between the entries and the body members.

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
            __p((Operation.ADD.apply(4, 2)).toString())
            __p((Operation.SUBTRACT.apply(7, 3)).toString())
            __p((Operation.MULTIPLY.apply(3, 5)).toString())
        
__check("6\n4\n15")
}
