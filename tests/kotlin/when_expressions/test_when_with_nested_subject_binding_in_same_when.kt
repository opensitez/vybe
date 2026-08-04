// vybe-test: kotlin/when_expressions/test_when_with_nested_subject_binding_in_same_when
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun decode(value: Any): String {
            return when (value) {
                is Int -> {
                    val doubled = value * 2
                    when {
                        doubled > 10 -> "int-big"
                        else -> "int-small"
                    }
                }
                is String -> {
                    val head = value.firstOrNull() ?: '?'
                    when (head) {
                        in 'a'..'m' -> "string-low"
                        in 'n'..'z' -> "string-high"
                        else -> "string-other"
                    }
                }
                else -> "none"
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
            __p((decode(7)).toString())
            __p((decode(4)).toString())
            __p((decode("beta")).toString())
            __p((decode("zeta")).toString())
            __p((decode("@")).toString())
            __p((decode(3.0)).toString())
        
__check("int-small\nint-small\nstring-low\nstring-high\nstring-other\nnone")
}
