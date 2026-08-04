// vybe-test: kotlin/operators/test_custom_index_get_set
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Buckets {
            private val data = arrayOf(5, 10, 15)
            operator fun get(index: Int): Int {
                return data[index]
            }
            operator fun set(index: Int, value: Int) {
                data[index] = value
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
            val storage = Buckets()
            __p((storage[0]).toString())
            storage[1] = 25
            __p((storage[1]).toString())
            __p((storage[2]).toString())
        
__check("5\n25\n15")
}
