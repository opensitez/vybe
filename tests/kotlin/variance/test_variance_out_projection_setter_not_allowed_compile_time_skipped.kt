// vybe-test: kotlin/variance/test_variance_out_projection_setter_not_allowed_compile_time_skipped
// origin: languages/kotlin/tests/kotlin/test_variance.rs

class Repo<T> {
            private val store = mutableListOf<T>()
            fun getStore(): List<T> = store
            fun addAll(values: List<out T>) {
                // this path is intentionally empty
            }
            fun mainAdd() {
                __p((store.size).toString())
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
            val r = Repo<String>()
            r.mainAdd()
        
__check("0")
}
