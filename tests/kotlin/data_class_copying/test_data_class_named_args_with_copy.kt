// vybe-test: kotlin/data_class_copying/test_data_class_named_args_with_copy
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Config(val host: String, val port: Int, val secure: Boolean)
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
            val base = Config("localhost", 80, false)
            val secure = base.copy(port = 443, secure = true)
            __p((secure.host).toString())
            __p((secure.port).toString())
            __p((secure.secure).toString())
        
__check("localhost\n443\ntrue")
}
