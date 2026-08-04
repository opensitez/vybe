// vybe-test: kotlin/scope/test_scope_after_smart_cast_branch_isolated_by_type
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun label(value: Any): String {
            return when (value) {
                is String -> "str:" + value.length
                is Int -> "int:" + value
                is Boolean -> "bool:" + value
                else -> "other"
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
            __p((label("abc")).toString())
            __p((label(9)).toString())
            __p((label(true)).toString())
            __p((label(2.5)).toString())
        
__check("str:3\nint:9\nbool:true\nother")
}
