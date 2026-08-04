// vybe-test: kotlin/when_expressions/test_when_on_data_class_property_subject
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

data class User(val name: String, val active: Boolean, val level: Int)

        fun label(user: User): String {
            return when {
                user.name.isEmpty() -> "anon"
                !user.active -> "inactive"
                user.level > 10 -> "vip"
                else -> "regular"
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
            __p((label(User("", true, 3))).toString())
            __p((label(User("a", false, 1))).toString())
            __p((label(User("b", true, 12))).toString())
            __p((label(User("c", true, 4))).toString())
        
__check("anon\ninactive\nvip\nregular")
}
